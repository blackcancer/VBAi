# Native renderer tests

These tests cover the lifecycle and loading of `VBAi.Native.dll`. They are not a
replacement for visual qualification in real VBE hosts and do not measure C++
line or branch coverage.

## Synthetic C++ cycle

From a Windows development environment with the Visual Studio C++ x64 tools and
Windows SDK, run at the repository root:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File tools/build/Build-NativeRenderer.ps1 -OutputDirectory artifacts/native-check -BuildSelfTest
& ./artifacts/native-check/NativeRendererSelfTest.exe "$PWD/artifacts/native-check/VBAi.Native.dll" "$PWD/artifacts/native-check/VBE7.dll"
```

The `VBE7.dll` in this output is a **synthetic test fixture**. Never install it into
an application or distribute it as the VBA runtime. The fixture lets the self-test
exercise repeated native lifecycle operations without hooking Office.

## Managed loader

Build the isolated solution output described in [testing](../README.md), then run:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File tests/native/Test-NativeRendererLoader.ps1 -AssemblyPath artifacts/build/VBAi/Debug/net48/VBAi.dll
```

The loader checks extraction, payload identity and the managed/native contract.
Successful loading does not establish correct painting, focus, DPI handling,
reentrancy or compatibility with a particular Office/VBA update.

## Real-host qualification

Keep native appearance work explicitly experimental. Use a disposable test session,
record the host/build and exercise activation, repaint, theme changes, docking,
multiple windows and shutdown. Restore the original palette and options. Never
terminate an unrelated process to make a visual test pass.

Consult [appearance recovery](../../docs/troubleshooting.md#appearance-and-native-palette-recovery)
and [recorded validation](../../docs/test-coverage.md) before claiming support for a
new host or renderer path.
