# Native renderer checks

Run from the repository root, with the Visual Studio C++ x64 workload installed:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File tools/build/Build-NativeRenderer.ps1 -OutputDirectory artifacts/native-check -BuildSelfTest
& ./artifacts/native-check/NativeRendererSelfTest.exe "$PWD/artifacts/native-check/CodexVBE.Native.dll" "$PWD/artifacts/native-check/VBE7.dll"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File tests/native/Test-NativeRendererLoader.ps1 -AssemblyPath artifacts/native-product-integration/build/CodexVBE/Debug/net48/CodexVBE.dll
```

The lifecycle executable creates invisible synthetic windows in its own process.
Its synthetic VBE7.dll must never be copied into an Office or installation folder.
The loader check requires a compiled managed add-in, validates its embedded native
payload and uses the versioned local cache; it does not activate hooks.
These checks do not open Office and do not establish visual compatibility.
