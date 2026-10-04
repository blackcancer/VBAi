# Source-build setup

The standalone installer is planned for a later milestone. This page describes
**developer registration of the current preview**, not an end-user installation
experience or an installer implementation task.

## Requirements

Use Windows and a compatible 64-bit VBE host. Install Visual Studio with the .NET
desktop development tools, the .NET Framework 4.8 targeting/developer tools, the
C++ x64 build tools and a Windows SDK. `TlbExp.exe` comes from the .NET Framework
SDK. The build invokes the native renderer build and requires those C++ tools.

The native Office adapters also need the `Microsoft.Office.Interop.PowerPoint`,
`Microsoft.Office.Interop.Word` and `Microsoft.Vbe.Interop` 15.0 PIAs at build time. The project searches
the Visual Studio Office tools and the GAC. If they are installed elsewhere, pass
`-p:PowerPointInteropPath="<path-to>/Microsoft.Office.Interop.PowerPoint.dll"`
`-p:WordInteropPath="<path-to>/Microsoft.Office.Interop.Word.dll"` and
`-p:VbeInteropPath="<path-to>/Microsoft.Vbe.Interop.dll"` to
`dotnet build`. `EmbedInteropTypes=true` embeds the types used by VBAi; these PIAs
are not separate files to deploy with the add-in.

Monaco requires the WebView2 Runtime. Git workflows additionally require Git for
Windows and suitable credentials. Provider requirements are separate; see
[providers](providers.md). Do not download prerequisites from untrusted mirrors.

## Build without replacing a loaded DLL

From the repository root in a 64-bit Visual Studio development PowerShell:

```powershell
dotnet build VBAi.sln -c Debug -p:BuildOutputRoot="$PWD/artifacts/build"
```

This produces an isolated output under `artifacts/build/`. It does not register
the add-in or replace the assembly currently loaded by an application. See
[development](development.md) for output conventions.

## Register a development build

Save your work and close every host that has loaded the add-in. The current
registration scripts expect the standard Debug output:

```powershell
dotnet build VBAi.sln -c Debug -p:Platform=x64
TlbExp.exe .\bin\Debug\net48\VBAi.dll /out:.\bin\Debug\net48\VBAi.tlb
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\Install-VBAi.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\Test-VBAiInstallation.ps1
```

Review scripts before running them. `ExecutionPolicy Bypass` above applies to the
invoked PowerShell process; it is not an instruction to change the machine's
permanent policy. Registration writes COM/VBE registry entries and may request
elevation for the machine-view ProgID resolution used by the native pane.

Keep the complete build output, including resources and dependencies. The scripts
check the type library, COM identity and registration; a missing or stale TLB is
not accepted. Do not copy only the managed DLL to another location after registration.

## Verify the loaded assembly

Open the VBE normally from the host and check the VBAi menu entries. For a known
host PID, inspect the bridge:

```powershell
powershell.exe -NoProfile -File .\tools\Invoke-VBAi.ps1 -HostProcessId 12345 -Command status
```

Replace `12345` with the intended process ID. Check the loaded assembly path and
build identity, not only whether a menu is visible. Close and restart the entire
host after replacing a build; hiding the VBE window can leave the assembly loaded.

The registered identities are `VBAi.AddIn` and `VBAi.ChatToolWindow`; the named pipe
is `VBAi.<PID>`. The VBE discovery location used by the tested setup is
`HKCU\Software\Microsoft\VBA\VBE\6.0\Addins64`. Do not assume that a registry entry
alone establishes successful loading in another host.

## Sandboxed development shells

When `CODEX_SHELL=1` is present, the scripts use
`Invoke-VBAi-OutsideSandbox.ps1` and a temporary scheduled task to perform and
verify registration in the real Windows registry. Inspect the result of that
operation; an isolated shell's registry view is not loading evidence.

## Legacy installations and removal

The VBAi rename retains COM GUIDs but updates names and ProgIDs. The registration
migration recognizes the known legacy identity. User-data migration copies to
VBAi locations without overwriting existing destination files; original data is
retained and reparse links are refused. Back up configuration and history before
manual maintenance.

To unregister, close the hosts and run `tools/Uninstall-VBAi.ps1`. Do not delete
history, recovery folders or unpushed Git objects as a generic cleanup step.
See [troubleshooting](troubleshooting.md) for diagnostic paths and
[updates](updates.md) for the future distribution contract.
