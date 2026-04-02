# DesktopBuddy Mod - Implementation Plan

## Overview
Eight changes across 3 modified files, 3 deleted files, and 1 updated doc file.

---

## 1. DesktopSession Class Changes (DesktopBuddyMod.cs lines 364-374)

Add three new fields to DesktopSession:

- Component ActiveSource -- Hand tracking: which interaction source owns this panel
- float ScrollAccumulator -- Joystick scroll: fractional scroll accumulation
- IntPtr WindowHandle -- Store hwnd for focus-on-spawn and other uses

Rationale: ActiveSource solves the two-hand fighting issue. ScrollAccumulator enables smooth joystick-to-discrete-scroll-click conversion. WindowHandle stores hwnd on the session object.

---

## 2. Desktop Icon on Context Menu (ContextMenuPatch.cs line 142)

### Current Code
Line 142 uses the Uri overload with null, producing no icon.

### Change
Generate a 32x32 procedural monitor icon as RGBA bytes, save to LocalDB, cache the URI, use the IAssetProvider<ITexture2D> overload.

### Implementation

Add a static method GetDesktopIcon(Engine engine, Slot slot) that:

1. Check a static Uri _desktopIconUri cache field. If non-null, create StaticTexture2D, set URL, return it.
2. If null, generate 32x32 RGBA byte array programmatically:
   - Fill all pixels with transparent (0,0,0,0)
   - Draw monitor body (cols 2-29, rows 2-22) in dark blue (40, 80, 140, 255)
   - Draw inner screen (cols 4-27, rows 4-20) in light blue (100, 160, 220, 255)
   - Draw stand: vertical bar (cols 14-17, rows 23-26) and base (cols 10-21, rows 27-28) in dark gray (80, 80, 80, 255)
3. Save bitmap to LocalDB async (same pattern as window icons).
4. Cache the URI in static field.
5. Create StaticTexture2D on the menu slot, set URL, return it.

### Modified Postfix
Use GetDesktopIcon to get an icon texture. If non-null, call menu.AddItem with the IAssetProvider<ITexture2D> overload. Otherwise fall back to the Uri overload.

---

## 3. VR Hand Hover Fighting (DesktopBuddyMod.cs)

### Problem
Both VR hands fire LocalHoverStay simultaneously, causing cursor teleportation.

### Solution
Track which Component source is the active controller per session via session.ActiveSource.

### Implementation

a) LocalPressed handler (line 116): Add session.ActiveSource = data.source at top.

b) LocalHoverStay handler (line 138): Add guard at start - if session.ActiveSource != null and data.source != session.ActiveSource, return immediately.

c) LocalHoverEnter handler (line 111): Add same guard.

ActiveSource is set on click and joystick scroll only. Not on hover alone, preventing accidental takeover.

---

## 4. VR Joystick Scroll (DesktopBuddyMod.cs)

### Current Code (lines 145-151)
Only reads mouse.ScrollWheelDelta.

### New Code
After SendHover, replace the scroll block with two parts:

Part 1 - Mouse scroll (desktop mode): Keep existing logic reading mouse.ScrollWheelDelta.

Part 2 - VR joystick scroll:
1. Get InteractionHandler via data.source.Slot.GetComponentInParents<InteractionHandler>()
2. Read handler.Side.Value to get Chirality
3. Get controller via root.World.InputInterface.GetControllerNode(side)
4. Read controller.Axis.Value.y for vertical joystick
5. Apply 0.15f deadzone (matches engine DEFAULT_MOVEMENT_THRESHOLD)
6. Accumulate: scrollDelta = joy * Time.Delta * 5f (about 5 clicks/sec at full deflection)
7. Add to session.ScrollAccumulator, send integer clicks when abs >= 1.0
8. Reset accumulator when in deadzone
9. Claim session.ActiveSource on joystick scroll
10. Wrap in try/catch for safety

---

## 5. Keyboard Button Fix (DesktopBuddyMod.cs lines 154-174)

### Root Cause
SimpleVirtualKeyboard.OnAttach() calls new UIBuilder(base.Slot, 640f, 192f, 0.001f) which sets keyboardSlot.LocalScale = 0.001. Root already has 0.001 scale. Result: 0.001 squared = 0.000001 effective scale. Keyboard is invisible.

### Fix
After AttachComponent<SimpleVirtualKeyboard>(), reset keyboardSlot.LocalScale = float3.One.

This works because OnAttach runs synchronously during AttachComponent. The UIBuilder inside OnAttach sets scale to 0.001, then our next line resets it to 1.0. With root at 0.001, effective scale is 0.001 * 1.0 = 0.001 (correct).

Keyboard stays parented to root so it moves with the panel when grabbed.

Fallback if OnAttach is async (unlikely): root.World.RunInUpdates(1, () => { keyboardSlot.LocalScale = float3.One; })

---

## 6. Focus Window on Spawn (DesktopBuddyMod.cs)

After ActiveSessions.Add(session) (line 188), add:
- WindowInput.FocusWindow(hwnd)
- session.WindowHandle = hwnd

---

## 7. Enable Scaling (DesktopBuddyMod.cs line 176)

Change Scalable.Value from false to true.

Safe because normalizedPressPoint is Canvas-local (0-1 regardless of world scale), and streamer.Width/Height are Win32 pixel values unaffected by in-game scale.

---

## 8. Dead Code Removal

### Files to Delete
1. DesktopBuddy/WindowCapture.cs - Old GDI capture, replaced by WgcCapture. Zero refs.
2. DesktopBuddy/MjpegServer.cs - FFmpeg server. Zero code refs.
3. DesktopBuddy/PopupWatcher.cs - Popup hook watcher. Zero refs.

### Dead Method
Remove HandleTouch (lines 333-358) - private, zero callers.

### Comment Cleanup
Remove lines 24-25 (MjpegServer comment).

---

## 9. Logging

Add Msg() at: hover guard (throttled), hover-enter guard, active source claim on press, active source claim on joystick, joystick scroll send (throttled), focus on spawn, scale in click, keyboard creation. Throttle high-frequency with _updateCount % 300 == 0.

---

## 10. CLAUDE.md Updates

Update Source Files list (remove WindowCapture/MjpegServer, add WgcCapture/WindowIconExtractor). Add Logging Convention section.

---

## Implementation Order

1. Dead code removal
2. DesktopSession fields
3. Focus on spawn
4. Enable scaling
5. Hand hover fighting
6. Joystick scroll
7. Keyboard fix
8. Desktop icon
9. Logging pass
10. CLAUDE.md update

---

## Risk Assessment

| Change | Risk | Mitigation |
|--------|------|------------|
| Hand tracking | Low | Falls back if ActiveSource is null |
| Joystick scroll | Low | try/catch; accumulator reset |
| Keyboard fix | Medium | Relies on OnAttach sync. Fallback: RunInUpdates |
| Scaling | Low | normalizedPressPoint is scale-independent |
| Desktop icon | Low | Graceful no-icon fallback |
| Dead code | Very Low | Verified zero references |
| Focus on spawn | Very Low | Already used elsewhere |
