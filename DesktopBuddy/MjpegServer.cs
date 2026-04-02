using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Text;
using System.Threading;
using ResoniteModLoader;

namespace DesktopBuddy;

/// <summary>
/// HTTP server that serves MPEG-TS video streams encoded from WGC capture frames via FFmpeg.
/// Frames come from DesktopSession.StreamFrame (set by the update loop, one pointer assignment).
/// A feeder thread reads frames at 30fps and pipes raw RGBA to FFmpeg stdin.
/// FFmpeg encodes to H.264 MPEG-TS on stdout, served to HTTP clients.
///
/// Endpoints:
///   GET /stream?session={index}  → MPEG-TS stream for a specific session (0-based index into ActiveSessions)
/// </summary>
public sealed class MjpegServer : IDisposable
{
    private HttpListener _listener;
    private readonly Thread _listenThread;
    private volatile bool _running;
    private readonly int _port;

    private readonly List<Process> _ffmpegProcesses = new();
    private readonly object _processLock = new();

    // Track active stream per session to kill old one on retry
    private readonly Dictionary<int, Process> _activeStreams = new();

    public int Port => _port;

    public MjpegServer(int port = 48080)
    {
        _port = port;
        _listener = new HttpListener();
        _running = true;
        _listenThread = new Thread(ListenLoop) { IsBackground = true, Name = "DesktopBuddy_HTTP" };
    }

    public void Start()
    {
        try
        {
            _listener.Prefixes.Add($"http://+:{_port}/");
            _listener.Start();
            ResoniteMod.Msg($"[MjpegServer] Listening on http://+:{_port}/");
        }
        catch
        {
            _listener.Close();
            _listener = new HttpListener();
            _listener.Prefixes.Add($"http://localhost:{_port}/");
            _listener.Start();
            ResoniteMod.Msg($"[MjpegServer] Listening on http://localhost:{_port}/ (run 'netsh http add urlacl url=http://+:48080/ user=Everyone' as admin for tunnel)");
        }
        _listenThread.Start();
    }

    private void ListenLoop()
    {
        while (_running)
        {
            try
            {
                var ctx = _listener.GetContext();
                ThreadPool.QueueUserWorkItem(_ => HandleRequest(ctx));
            }
            catch (HttpListenerException) { break; }
            catch { }
        }
    }

    private void HandleRequest(HttpListenerContext ctx)
    {
        try
        {
            var path = ctx.Request.Url?.AbsolutePath ?? "/";
            if (path.StartsWith("/stream"))
                ServeStream(ctx);
            else
            {
                ctx.Response.StatusCode = 404;
                ctx.Response.Close();
            }
        }
        catch { try { ctx.Response.Close(); } catch { } }
    }

    private void ServeStream(HttpListenerContext ctx)
    {
        var remoteAddr = ctx.Request.RemoteEndPoint;
        var hostHeader = ctx.Request.Headers["Host"];
        var rawUrl = ctx.Request.RawUrl;
        ResoniteMod.Msg($"[MjpegServer] Stream request from {remoteAddr} Host={hostHeader} URL={rawUrl}");

        var sessionStr = ctx.Request.QueryString["session"];
        int sessionIdx = 0;
        if (sessionStr != null) int.TryParse(sessionStr, out sessionIdx);

        if (sessionIdx < 0 || sessionIdx >= DesktopBuddyMod.ActiveSessions.Count)
        {
            ResoniteMod.Msg($"[MjpegServer] Session {sessionIdx} not found, {DesktopBuddyMod.ActiveSessions.Count} active");
            ctx.Response.StatusCode = 404;
            ctx.Response.Close();
            return;
        }

        var session = DesktopBuddyMod.ActiveSessions[sessionIdx];
        int w = session.StreamWidth > 0 ? session.StreamWidth : 1920;
        int h = session.StreamHeight > 0 ? session.StreamHeight : 1080;

        // Kill any existing stream for this session (VideoTextureProvider retries cause duplicates)
        lock (_processLock)
        {
            if (_activeStreams.TryGetValue(sessionIdx, out var oldProc))
            {
                ResoniteMod.Msg($"[MjpegServer] Killing old FFmpeg for session {sessionIdx} PID={oldProc.Id}");
                try { oldProc.Kill(); } catch { }
                _activeStreams.Remove(sessionIdx);
            }
        }

        var ffmpegPath = FindFfmpeg();
        if (ffmpegPath == null)
        {
            ResoniteMod.Msg($"[MjpegServer] FFmpeg not found");
            ctx.Response.StatusCode = 500;
            ctx.Response.Close();
            return;
        }

        string encoder = DetectEncoder(ffmpegPath);

        // Build encoder options
        string encoderOpts;
        if (encoder == "h264_nvenc")
            encoderOpts = $"-c:v {encoder} -preset p1 -tune ll -rc vbr -cq 23 -b:v 8M -maxrate 12M";
        else if (encoder == "h264_qsv")
            encoderOpts = $"-c:v {encoder} -global_quality 23";
        else if (encoder == "h264_amf")
            encoderOpts = $"-c:v {encoder} -quality speed -rc cqp -qp_i 23 -qp_p 23";
        else
            encoderOpts = $"-c:v {encoder} -preset ultrafast -tune zerolatency -crf 23";

        // FFmpeg reads raw RGBA from stdin, encodes to H.264 MPEG-TS on stdout
        // vflip: WGC buffer is Y-flipped for GPU upload, need to flip back for video
        // scale: NVENC H.264 max width is 4096
        string scaleOpt = (encoder == "h264_nvenc" && w > 4096) ? ",scale=4096:-2" : "";
        string args = $"-f rawvideo -pixel_format rgba -video_size {w}x{h} -framerate 30 -i pipe:0 " +
                      $"-vf \"vflip{scaleOpt}\" -pix_fmt yuv420p {encoderOpts} -bf 0 -g 15 " +
                      $"-f mpegts -an pipe:1";

        ResoniteMod.Msg($"[MjpegServer] Starting FFmpeg: {ffmpegPath} {args}");

        var psi = new ProcessStartInfo
        {
            FileName = ffmpegPath,
            Arguments = args,
            UseShellExecute = false,
            RedirectStandardInput = true,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true,
        };

        Process ffmpeg;
        try
        {
            ffmpeg = Process.Start(psi)!;
            ResoniteMod.Msg($"[MjpegServer] FFmpeg started PID={ffmpeg.Id} for session {sessionIdx} ({w}x{h})");
        }
        catch (Exception ex)
        {
            ResoniteMod.Msg($"[MjpegServer] FFmpeg start FAILED: {ex.Message}");
            ctx.Response.StatusCode = 500;
            ctx.Response.Close();
            return;
        }

        ffmpeg.ErrorDataReceived += (s, e) =>
        {
            if (e.Data != null)
                ResoniteMod.Msg($"[FFmpeg:{ffmpeg.Id}] {e.Data}");
        };
        ffmpeg.BeginErrorReadLine();
        lock (_processLock)
        {
            _ffmpegProcesses.Add(ffmpeg);
            _activeStreams[sessionIdx] = ffmpeg;
        }

        // Feeder thread: reads WGC frames from session, writes to FFmpeg stdin at 30fps
        var feederThread = new Thread(() =>
        {
            try
            {
                var stdin = ffmpeg.StandardInput.BaseStream;
                int frameCount = 0;
                while (_running && !ffmpeg.HasExited && session.Root != null && !session.Root.IsDestroyed)
                {
                    byte[] frame = null;
                    int fw = 0, fh = 0;
                    lock (session.StreamLock)
                    {
                        // Wait for a new frame (up to 100ms)
                        if (!session.StreamFrameReady)
                            Monitor.Wait(session.StreamLock, 100);
                        if (session.StreamFrameReady)
                        {
                            frame = session.StreamFrame;
                            fw = session.StreamWidth;
                            fh = session.StreamHeight;
                            session.StreamFrameReady = false;
                        }
                    }
                    if (frame != null && fw == w && fh == h)
                    {
                        try
                        {
                            stdin.Write(frame, 0, fw * fh * 4);
                            stdin.Flush();
                            frameCount++;
                            if (frameCount <= 3 || frameCount % 300 == 0)
                                ResoniteMod.Msg($"[MjpegServer] Fed frame #{frameCount} to FFmpeg ({fw}x{fh})");
                        }
                        catch { break; }
                    }
                    else if (frame != null && (fw != w || fh != h))
                    {
                        // Resolution changed — need to restart FFmpeg
                        ResoniteMod.Msg($"[MjpegServer] Resolution changed {w}x{h} -> {fw}x{fh}, ending stream");
                        break;
                    }
                }
                try { stdin.Close(); } catch { }
                ResoniteMod.Msg($"[MjpegServer] Feeder done, fed {frameCount} frames");
            }
            catch (Exception ex)
            {
                ResoniteMod.Msg($"[MjpegServer] Feeder error: {ex.Message}");
            }
        }) { IsBackground = true, Name = $"DesktopBuddy_Feeder_{sessionIdx}" };
        feederThread.Start();

        // Serve FFmpeg stdout to HTTP response
        ctx.Response.ContentType = "video/mp2t";
        ctx.Response.SendChunked = true;
        ctx.Response.StatusCode = 200;

        long totalBytes = 0;
        int readCount = 0;
        try
        {
            var buffer = new byte[65536];
            var stdout = ffmpeg.StandardOutput.BaseStream;
            ResoniteMod.Msg($"[MjpegServer] Serving stream to client...");
            while (_running && !ffmpeg.HasExited)
            {
                int read = stdout.Read(buffer, 0, buffer.Length);
                if (read <= 0) break;
                ctx.Response.OutputStream.Write(buffer, 0, read);
                ctx.Response.OutputStream.Flush();
                totalBytes += read;
                readCount++;
                if (readCount <= 3 || readCount % 300 == 0)
                    ResoniteMod.Msg($"[MjpegServer] Sent chunk #{readCount}: {read} bytes (total: {totalBytes})");
            }
        }
        catch (Exception ex)
        {
            ResoniteMod.Msg($"[MjpegServer] Stream exception: {ex.GetType().Name}: {ex.Message}");
        }
        finally
        {
            try { if (!ffmpeg.HasExited) ffmpeg.Kill(); } catch { }
            lock (_processLock)
            {
                _ffmpegProcesses.Remove(ffmpeg);
                if (_activeStreams.TryGetValue(sessionIdx, out var cur) && cur == ffmpeg)
                    _activeStreams.Remove(sessionIdx);
            }
            try { ctx.Response.Close(); } catch { }
            ResoniteMod.Msg($"[MjpegServer] Stream ended session={sessionIdx}. Total: {totalBytes} bytes, {readCount} chunks");
        }
    }

    private static string? _cachedEncoder;

    private static string DetectEncoder(string ffmpegPath)
    {
        if (_cachedEncoder != null) return _cachedEncoder;
        string[] encoders = { "h264_nvenc", "h264_amf", "h264_qsv", "libx264" };
        foreach (var enc in encoders)
        {
            try
            {
                var p = Process.Start(new ProcessStartInfo
                {
                    FileName = ffmpegPath,
                    Arguments = $"-f lavfi -i nullsrc=s=256x256:d=0.1 -c:v {enc} -f null -",
                    RedirectStandardOutput = true, RedirectStandardError = true,
                    UseShellExecute = false, CreateNoWindow = true
                });
                p?.WaitForExit(3000);
                if (p?.ExitCode == 0) { _cachedEncoder = enc; return enc; }
            }
            catch { }
        }
        _cachedEncoder = "libx264";
        return "libx264";
    }

    private static string? FindFfmpeg()
    {
        string[] candidates = {
            @"C:\bins\ffmpeg.exe",
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "ffmpeg", "bin", "ffmpeg.exe"),
            "ffmpeg"
        };
        foreach (var c in candidates)
        {
            try
            {
                var p = Process.Start(new ProcessStartInfo { FileName = c, Arguments = "-version", RedirectStandardOutput = true, UseShellExecute = false, CreateNoWindow = true });
                p?.WaitForExit(1000);
                if (p?.ExitCode == 0) return c;
            }
            catch { }
        }
        return null;
    }

    public void Dispose()
    {
        _running = false;
        lock (_processLock)
        {
            foreach (var p in _ffmpegProcesses)
            {
                try { p.Kill(); } catch { }
            }
            _ffmpegProcesses.Clear();
        }
        try { _listener.Stop(); } catch { }
        try { _listener.Close(); } catch { }
    }
}
