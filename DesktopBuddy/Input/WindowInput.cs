using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using ResoniteModLoader;

namespace DesktopBuddy;

/// <summary>
/// Handles Windows input injection for the desktop viewer.
/// Touch injection follows the official Microsoft sample:
/// https://github.com/microsoft/Windows-classic-samples/blob/main/Samples/InjectTouchPointerInput/cpp/InjectTouch.cpp
/// </summary>
public static class WindowInput
{
    // --- Mouse APIs (hover + scroll) ---

    [DllImport("user32.dll")]
    private static extern bool ClientToScreen(IntPtr hWnd, ref POINT lpPoint);

    [DllImport("user32.dll")]
    private static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern bool SetCursorPos(int X, int Y);

    [DllImport("user32.dll")]
    private static extern void mouse_event(uint dwFlags, int dx, int dy, int dwData, IntPtr dwExtraInfo);

    private const uint MOUSEEVENTF_WHEEL = 0x0800;

    // --- Touch injection APIs ---
    // Exact signatures from MSDN winuser.h

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool InitializeTouchInjection(uint maxCount, uint dwMode);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool InjectTouchInput(uint count, [In] POINTER_TOUCH_INFO[] contacts);

    // InitializeTouchInjection dwMode — Microsoft sample uses TOUCH_FEEDBACK_NONE
    private const uint TOUCH_FEEDBACK_NONE = 0x3;

    // POINTER_FLAGS — exact hex from MSDN "Pointer Flags" reference
    // Microsoft sample uses ONLY these three for down:
    //   POINTER_FLAG_DOWN | POINTER_FLAG_INRANGE | POINTER_FLAG_INCONTACT
    // And ONLY POINTER_FLAG_UP for up.
    private const uint POINTER_FLAG_INRANGE    = 0x00000002;
    private const uint POINTER_FLAG_INCONTACT  = 0x00000004;
    private const uint POINTER_FLAG_DOWN       = 0x00010000;
    private const uint POINTER_FLAG_UPDATE     = 0x00020000;
    private const uint POINTER_FLAG_UP         = 0x00040000;

    // POINTER_INPUT_TYPE
    private const uint PT_TOUCH = 0x00000002;

    // TOUCH_FLAGS — Microsoft sample uses TOUCH_FLAG_NONE (0)
    private const uint TOUCH_FLAG_NONE = 0x00000000;

    // TOUCH_MASK — Microsoft sample uses all three
    private const uint TOUCH_MASK_CONTACTAREA = 0x00000004;
    private const uint TOUCH_MASK_ORIENTATION = 0x00000008;
    private const uint TOUCH_MASK_PRESSURE    = 0x00000002;

    [StructLayout(LayoutKind.Sequential)]
    public struct POINT { public int X, Y; }

    // POINTER_INFO — exact field order from MSDN ns-winuser-pointer_info
    // Microsoft sample zeroes entire struct with memset, then sets only what's needed
    [StructLayout(LayoutKind.Sequential)]
    private struct POINTER_INFO
    {
        public uint pointerType;
        public uint pointerId;
        public uint frameId;
        public uint pointerFlags;
        public IntPtr sourceDevice;
        public IntPtr hwndTarget;
        public POINT ptPixelLocation;
        public POINT ptHimetricLocation;
        public POINT ptPixelLocationRaw;
        public POINT ptHimetricLocationRaw;
        public uint dwTime;
        public uint historyCount;
        public int inputData;
        public uint dwKeyStates;
        public ulong PerformanceCount;
        public int ButtonChangeType;
    }

    // POINTER_TOUCH_INFO — exact field order from MSDN ns-winuser-pointer_touch_info
    [StructLayout(LayoutKind.Sequential)]
    private struct POINTER_TOUCH_INFO
    {
        public POINTER_INFO pointerInfo;
        public uint touchFlags;
        public uint touchMask;
        public RECT rcContact;
        public RECT rcContactRaw;
        public uint orientation;
        public uint pressure;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct RECT
    {
        public int Left, Top, Right, Bottom;
    }

    // --- State ---

    private static bool _touchInitialized;
    private static readonly object _initLock = new();
    private const uint MAX_TOUCH_CONTACTS = 10;

    // Track last position per touch ID (MSDN: UP position should match last frame)
    private static readonly POINT[] _lastPosition = new POINT[MAX_TOUCH_CONTACTS];
    // Track if move logging has been done (to avoid spam)
    private static readonly bool[] _moveFailLogged = new bool[MAX_TOUCH_CONTACTS];

    private static bool EnsureTouchInit()
    {
        if (_touchInitialized) return true;
        lock (_initLock)
        {
            if (_touchInitialized) return true;
            // Microsoft sample: InitializeTouchInjection(10, TOUCH_FEEDBACK_NONE)
            bool ok = InitializeTouchInjection(MAX_TOUCH_CONTACTS, TOUCH_FEEDBACK_NONE);
            if (!ok)
            {
                int err = Marshal.GetLastWin32Error();
                ResoniteMod.Msg($"[WindowInput] InitializeTouchInjection FAILED err={err}");
                return false;
            }
            _touchInitialized = true;
            ResoniteMod.Msg($"[WindowInput] Touch injection initialized (max={MAX_TOUCH_CONTACTS}, feedback=NONE)");
            return true;
        }
    }

    [DllImport("user32.dll")]
    private static extern bool IsIconic(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

    private const int SW_RESTORE = 9;

    /// <summary>
    /// Restore a minimized window before capturing/interacting with it.
    /// </summary>
    public static void RestoreIfMinimized(IntPtr hWnd)
    {
        if (hWnd == IntPtr.Zero) return;
        if (IsIconic(hWnd))
        {
            ResoniteMod.Msg($"[WindowInput] Restoring minimized window hwnd={hWnd}");
            ShowWindow(hWnd, SW_RESTORE);
        }
    }

    // --- Coordinate conversion ---

    private static POINT UvToScreen(IntPtr hWnd, float u, float v, int clientW, int clientH)
    {
        int px = (int)(u * clientW);
        int py = (int)(v * clientH);
        var pt = new POINT { X = px, Y = py };
        if (hWnd != IntPtr.Zero)
            ClientToScreen(hWnd, ref pt);
        return pt;
    }

    // --- Public API: hover (mouse) ---

    public static void FocusWindow(IntPtr hWnd)
    {
        if (hWnd != IntPtr.Zero)
            SetForegroundWindow(hWnd);
    }

    public static void SendHover(IntPtr hWnd, float u, float v, int clientW, int clientH)
    {
        var pt = UvToScreen(hWnd, u, v, clientW, clientH);
        SetCursorPos(pt.X, pt.Y);
    }

    // --- Public API: touch ---
    // Follows Microsoft official sample EXACTLY:
    // Down:   pointerFlags = POINTER_FLAG_DOWN | POINTER_FLAG_INRANGE | POINTER_FLAG_INCONTACT
    // Move:   pointerFlags = POINTER_FLAG_UPDATE | POINTER_FLAG_INRANGE | POINTER_FLAG_INCONTACT
    // Up:     pointerFlags = POINTER_FLAG_UP
    // All use: touchMask = TOUCH_MASK_CONTACTAREA | TOUCH_MASK_ORIENTATION | TOUCH_MASK_PRESSURE
    //          orientation = 90, pressure = 32000

    public static void SendTouchDown(IntPtr hWnd, float u, float v, int clientW, int clientH, uint touchId = 0)
    {
        if (!EnsureTouchInit()) return;
        var pt = UvToScreen(hWnd, u, v, clientW, clientH);
        if (touchId < MAX_TOUCH_CONTACTS)
        {
            _lastPosition[touchId] = pt;
            _moveFailLogged[touchId] = false;
        }

        // Match Microsoft sample: memset to zero, then set only needed fields
        var contact = new POINTER_TOUCH_INFO();
        contact.pointerInfo.pointerType = PT_TOUCH;
        contact.pointerInfo.pointerId = touchId;
        contact.pointerInfo.ptPixelLocation = pt;
        contact.pointerInfo.pointerFlags = POINTER_FLAG_DOWN | POINTER_FLAG_INRANGE | POINTER_FLAG_INCONTACT;
        contact.touchFlags = TOUCH_FLAG_NONE;
        contact.touchMask = TOUCH_MASK_CONTACTAREA | TOUCH_MASK_ORIENTATION | TOUCH_MASK_PRESSURE;
        contact.orientation = 90;
        contact.pressure = 32000;
        contact.rcContact.Top = pt.Y - 2;
        contact.rcContact.Bottom = pt.Y + 2;
        contact.rcContact.Left = pt.X - 2;
        contact.rcContact.Right = pt.X + 2;

        var arr = new POINTER_TOUCH_INFO[] { contact };
        if (!InjectTouchInput(1, arr))
        {
            int err = Marshal.GetLastWin32Error();
            ResoniteMod.Msg($"[Touch] Down FAILED id={touchId} screen=({pt.X},{pt.Y}) err={err}");
        }
        else
        {
            ResoniteMod.Msg($"[Touch] Down OK id={touchId} screen=({pt.X},{pt.Y})");
        }
    }

    public static void SendTouchMove(IntPtr hWnd, float u, float v, int clientW, int clientH, uint touchId = 0)
    {
        if (!_touchInitialized) return;
        var pt = UvToScreen(hWnd, u, v, clientW, clientH);
        if (touchId < MAX_TOUCH_CONTACTS) _lastPosition[touchId] = pt;

        var contact = new POINTER_TOUCH_INFO();
        contact.pointerInfo.pointerType = PT_TOUCH;
        contact.pointerInfo.pointerId = touchId;
        contact.pointerInfo.ptPixelLocation = pt;
        contact.pointerInfo.pointerFlags = POINTER_FLAG_UPDATE | POINTER_FLAG_INRANGE | POINTER_FLAG_INCONTACT;
        contact.touchFlags = TOUCH_FLAG_NONE;
        contact.touchMask = TOUCH_MASK_CONTACTAREA | TOUCH_MASK_ORIENTATION | TOUCH_MASK_PRESSURE;
        contact.orientation = 90;
        contact.pressure = 32000;
        contact.rcContact.Top = pt.Y - 2;
        contact.rcContact.Bottom = pt.Y + 2;
        contact.rcContact.Left = pt.X - 2;
        contact.rcContact.Right = pt.X + 2;

        var arr = new POINTER_TOUCH_INFO[] { contact };
        if (!InjectTouchInput(1, arr))
        {
            if (touchId < MAX_TOUCH_CONTACTS && !_moveFailLogged[touchId])
            {
                _moveFailLogged[touchId] = true;
                int err = Marshal.GetLastWin32Error();
                ResoniteMod.Msg($"[Touch] Move FAILED id={touchId} err={err} (further move errors suppressed)");
            }
        }
    }

    public static void SendTouchUp(IntPtr hWnd, float u, float v, int clientW, int clientH, uint touchId = 0)
    {
        if (!_touchInitialized) return;
        // Microsoft sample: just POINTER_FLAG_UP, same position as last frame
        var pt = (touchId < MAX_TOUCH_CONTACTS) ? _lastPosition[touchId] : UvToScreen(hWnd, u, v, clientW, clientH);

        var contact = new POINTER_TOUCH_INFO();
        contact.pointerInfo.pointerType = PT_TOUCH;
        contact.pointerInfo.pointerId = touchId;
        contact.pointerInfo.ptPixelLocation = pt;
        contact.pointerInfo.pointerFlags = POINTER_FLAG_UP;

        var arr = new POINTER_TOUCH_INFO[] { contact };
        if (!InjectTouchInput(1, arr))
        {
            int err = Marshal.GetLastWin32Error();
            ResoniteMod.Msg($"[Touch] Up FAILED id={touchId} err={err}");
        }
        else
        {
            ResoniteMod.Msg($"[Touch] Up OK id={touchId}");
        }
    }

    // --- Public API: scroll (mouse wheel) ---

    /// <summary>
    /// Send mouse wheel scroll. wheelDelta is raw WHEEL_DELTA units (120 = one notch).
    /// </summary>
    public static void SendScroll(IntPtr hWnd, float u, float v, int clientW, int clientH, int wheelDelta)
    {
        var pt = UvToScreen(hWnd, u, v, clientW, clientH);
        SetCursorPos(pt.X, pt.Y);
        mouse_event(MOUSEEVENTF_WHEEL, 0, 0, wheelDelta, IntPtr.Zero);
    }

    // --- Keyboard injection via SendInput ---

    [DllImport("user32.dll", SetLastError = true)]
    private static extern uint SendInput(uint nInputs, INPUT[] pInputs, int cbSize);

    private const uint INPUT_KEYBOARD = 1;
    private const uint KEYEVENTF_KEYUP = 0x0002;
    private const uint KEYEVENTF_UNICODE = 0x0004;

    // INPUT struct: DWORD type + union(MOUSEINPUT|KEYBDINPUT|HARDWAREINPUT)
    // Union must be 32 bytes on x64 (MOUSEINPUT is the largest member)
    // Total INPUT size on x64: 4(type) + 4(padding) + 32(union) = 40 bytes
    [StructLayout(LayoutKind.Sequential)]
    private struct INPUT
    {
        public uint type;
        public INPUTUNION u;
    }

    // Force union to 40 bytes to match MOUSEINPUT size on x64
    // MOUSEINPUT: dx(4) + dy(4) + mouseData(4) + dwFlags(4) + time(4) + pad(4) + dwExtraInfo(8) = 32
    // We only use KEYBDINPUT (24 bytes) but must pad to 32 for the union
    [StructLayout(LayoutKind.Explicit, Size = 32)]
    private struct INPUTUNION
    {
        [FieldOffset(0)] public KEYBDINPUT ki;
    }

    // KEYBDINPUT: wVk(2) + wScan(2) + dwFlags(4) + time(4) + pad(4) + dwExtraInfo(8) = 24 on x64
    [StructLayout(LayoutKind.Sequential)]
    private struct KEYBDINPUT
    {
        public ushort wVk;
        public ushort wScan;
        public uint dwFlags;
        public uint time;
        public IntPtr dwExtraInfo;
    }

    /// <summary>
    /// Type a unicode character string to the focused window via SendInput KEYEVENTF_UNICODE.
    /// </summary>
    public static void SendString(string text)
    {
        if (string.IsNullOrEmpty(text)) return;
        ResoniteMod.Msg($"[Keyboard] SendString: \"{text}\"");
        var inputs = new INPUT[text.Length * 2]; // down + up for each char
        int idx = 0;
        foreach (char c in text)
        {
            inputs[idx++] = new INPUT
            {
                type = INPUT_KEYBOARD,
                u = new INPUTUNION
                {
                    ki = new KEYBDINPUT
                    {
                        wVk = 0,
                        wScan = c,
                        dwFlags = KEYEVENTF_UNICODE,
                    }
                }
            };
            inputs[idx++] = new INPUT
            {
                type = INPUT_KEYBOARD,
                u = new INPUTUNION
                {
                    ki = new KEYBDINPUT
                    {
                        wVk = 0,
                        wScan = c,
                        dwFlags = KEYEVENTF_UNICODE | KEYEVENTF_KEYUP,
                    }
                }
            };
        }
        uint sent = SendInput((uint)inputs.Length, inputs, Marshal.SizeOf<INPUT>());
        if (sent != inputs.Length)
        {
            int err = Marshal.GetLastWin32Error();
            ResoniteMod.Msg($"[Keyboard] SendString FAILED sent={sent}/{inputs.Length} err={err}");
        }
    }

    /// <summary>
    /// Send a virtual key press (down + up) to the focused window.
    /// </summary>
    public static void SendVirtualKey(ushort vk)
    {
        var inputs = new INPUT[]
        {
            new INPUT { type = INPUT_KEYBOARD, u = new INPUTUNION { ki = new KEYBDINPUT { wVk = vk } } },
            new INPUT { type = INPUT_KEYBOARD, u = new INPUTUNION { ki = new KEYBDINPUT { wVk = vk, dwFlags = KEYEVENTF_KEYUP } } },
        };
        uint sent = SendInput((uint)inputs.Length, inputs, Marshal.SizeOf<INPUT>());
        if (sent != inputs.Length)
        {
            int err = Marshal.GetLastWin32Error();
            ResoniteMod.Msg($"[Keyboard] SendVirtualKey FAILED vk=0x{vk:X2} sent={sent}/{inputs.Length} err={err}");
        }
    }

    // Track held modifier keys
    private static readonly HashSet<ushort> _heldModifiers = new();

    /// <summary>
    /// Send a virtual key DOWN only (no up). Used for modifier keys (Ctrl, Shift, Alt).
    /// </summary>
    public static void SendVirtualKeyDown(ushort vk)
    {
        if (_heldModifiers.Contains(vk)) return; // Already held
        _heldModifiers.Add(vk);
        var inputs = new INPUT[]
        {
            new INPUT { type = INPUT_KEYBOARD, u = new INPUTUNION { ki = new KEYBDINPUT { wVk = vk } } },
        };
        SendInput(1, inputs, Marshal.SizeOf<INPUT>());
    }

    /// <summary>
    /// Send Ctrl+V (paste) via SendInput.
    /// </summary>
    public static void SendPaste()
    {
        ResoniteMod.Msg("[Keyboard] Sending Ctrl+V (paste)");
        const ushort VK_CONTROL = 0xA2; // Left Ctrl
        const ushort VK_V = 0x56;
        var inputs = new INPUT[]
        {
            new INPUT { type = INPUT_KEYBOARD, u = new INPUTUNION { ki = new KEYBDINPUT { wVk = VK_CONTROL } } },
            new INPUT { type = INPUT_KEYBOARD, u = new INPUTUNION { ki = new KEYBDINPUT { wVk = VK_V } } },
            new INPUT { type = INPUT_KEYBOARD, u = new INPUTUNION { ki = new KEYBDINPUT { wVk = VK_V, dwFlags = KEYEVENTF_KEYUP } } },
            new INPUT { type = INPUT_KEYBOARD, u = new INPUTUNION { ki = new KEYBDINPUT { wVk = VK_CONTROL, dwFlags = KEYEVENTF_KEYUP } } },
        };
        uint sent = SendInput((uint)inputs.Length, inputs, Marshal.SizeOf<INPUT>());
        if (sent != inputs.Length)
            ResoniteMod.Msg($"[Keyboard] SendPaste FAILED sent={sent}/{inputs.Length} err={Marshal.GetLastWin32Error()}");
    }

    /// <summary>
    /// Release all held modifier keys.
    /// </summary>
    public static void ReleaseAllModifiers()
    {
        if (_heldModifiers.Count == 0) return;
        var inputs = new INPUT[_heldModifiers.Count];
        int i = 0;
        foreach (var vk in _heldModifiers)
        {
            inputs[i++] = new INPUT { type = INPUT_KEYBOARD, u = new INPUTUNION { ki = new KEYBDINPUT { wVk = vk, dwFlags = KEYEVENTF_KEYUP } } };
        }
        SendInput((uint)inputs.Length, inputs, Marshal.SizeOf<INPUT>());
        ResoniteMod.Msg($"[Keyboard] Released {_heldModifiers.Count} modifiers");
        _heldModifiers.Clear();
    }
}
