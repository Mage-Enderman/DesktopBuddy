# DesktopBuddy - Resonite Mod

## Project
Resonite mod that spawns a world-space desktop/window viewer with touch input. Uses context menu patch to add "Spawn Desktop" option.

## Build & Deploy
```
dotnet build DesktopBuddy/DesktopBuddy.csproj
# Copy to: C:\Program Files (x86)\Steam\steamapps\common\Resonite\rml_mods\
# Kill Resonite first if DLL locked: taskkill /F /IM Renderite.Host.exe
```

## Architecture
- No custom components (engine doesn't support mod component types properly)
- Uses `SolidColorTexture` (built-in ProceduralTexture) + `SetFromCurrentBitmap` for direct GPU texture upload
- Capture thread → frame buffer → `World.RunInUpdates` update loop → copy into shared-memory Bitmap2D → upload
- Context menu via Harmony patch on `ContextMenu.OpenMenu`
- Window picker via `UIBuilder` with `Button.LocalPressed` events (NOT synced delegates — lambdas crash synced delegates)

## CRITICAL: Always Verify Before Writing Code

**Before writing or modifying ANY code that uses FrooxEngine/Resonite APIs:**

1. **Look up every method** you intend to call — verify its exact parameter names, types, order, and return type in the decompiled sources.
2. **Look up every class/component** you reference — verify its inheritance chain, available fields, and sync members.
3. **Never assume** a method signature, field type, or API behavior from memory or pattern-matching. The engine has many similar-looking APIs with subtle differences.
4. **When modifying existing code**, read the surrounding code first and verify that the types/methods it uses still exist and have the signatures assumed.

This applies equally to new code and edits to existing code. A wrong assumption about a parameter type or method name will compile but crash at runtime in Resonite.

## CRITICAL: Log Everything

**Every code path must have `Msg()` logging.** If something doesn't work, we need maximum data to diagnose it. Log:
- Every event handler entry (with key parameter values)
- Every branch/decision point
- Every error with full exception details
- State transitions (source claimed, mode changed, etc.)

Never write a code path without a log statement. If it runs, we need to know it ran.

## Source Files
- `DesktopBuddyMod.cs` — Main mod, streaming setup, update loop, VR input handling
- `ContextMenuPatch.cs` — Harmony postfix on ContextMenu.OpenMenu, window picker, desktop icon
- `DesktopStreamer.cs` — Wraps WgcCapture, provides CaptureFrame(out w, out h)
- `WgcCapture.cs` — Windows.Graphics.Capture GPU-accelerated screen/window capture
- `WindowEnumerator.cs` — EnumWindows to list visible windows
- `WindowInput.cs` — Win32 mouse/scroll injection (SetCursorPos, mouse_event)
- `WindowIconExtractor.cs` — Extract window icons as RGBA via Win32
- `MjpegServer.cs` — HTTP server + FFmpeg gdigrab streaming (kept for future remote user support)

## Documentation Index

### API Reference (`docs/api-reference/`)
Summarized reference docs organized by topic. Use these for quick orientation on what's available.

| File | Topics |
|------|--------|
| 01 | Worker, Component, ComponentBase, ContainerWorker, WorkerInitializer |
| 02 | Sync data model: SyncField, SyncRef, SyncList, drivers, RefID |
| 03 | World, Engine, WorldManager, User, UserRoot, Userspace |
| 04 | Texture providers: AssetProvider chain, ProceduralTexture, StaticTexture2D, DesktopTextureProvider, VideoTextureProvider |
| 05 | Rendering: MeshRenderer, materials, meshes, lights, Camera, RenderSystem |
| 06 | UIX interaction: Canvas, RawImage, Image, Button, UIBuilder, InteractionElement, DesktopInteractionRelay |
| 07 | Physics: colliders, CharacterController, TouchSource, InteractionHandler, DevTool |
| 08 | Audio, animation, avatar, DynamicBoneChain |
| 09 | Tools, inspectors, utility, managers |
| 10 | Slot, SlotMeshes, TypeManager, GlobalTypeRegistry |
| 11 | Elements.Core: math, vectors, colors, animation, serialization, threading |
| 12 | Elements.Assets, SkyFrost.Base cloud API, FrooxEngine.Store, Commands |
| 13 | ProtoFlux runtime + 2575 core nodes + FrooxEngine nodes |
| 14 | Renderite.Shared IPC/render, Awwdio spatial audio/DSP |
| 15-27 | Full FrooxEngine.dll class-by-class (all 6209 types, chunked) |

### Decompiled Sources (`docs/decompiled_full/`)
Full decompiled source code — the ground truth. **Always verify method signatures here before writing code.**

| File | Contents |
|------|----------|
| `FrooxEngine.full.cs` | 585,342 lines, 6209 types — the entire engine |
| `Elements.Core.decompiled.cs` | 267,335 lines — math, color, vectors, animation, serialization |
| `Elements.Assets.decompiled.cs` | Bitmap2D, MeshX, AudioX, asset metadata |
| `SkyFrost.Base.decompiled.cs` | Cloud API, records, storage, variables |
| `FrooxEngine.Store.decompiled.cs` | LocalDB, asset records |
| `ProtoFlux.Core.decompiled.cs` | Execution runtime, node base classes |
| `ProtoFlux.Nodes.Core.decompiled.cs` | 2575 platform-independent nodes |
| `ProtoFlux.Nodes.FrooxEngine.decompiled.cs` | Resonite-specific nodes |
| `ProtoFluxBindings.decompiled.cs` | 22MB auto-generated bindings |
| `Renderite.Shared.decompiled.cs` | IPC, shared memory, render commands |
| `Awwdio.decompiled.cs` | Audio simulation, spatial audio, DSP |

### Individual Type Files (`docs/decompiled/`)
133 individually decompiled key types — useful for focused lookup of specific classes.
