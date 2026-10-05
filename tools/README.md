# Development tools

Use the command that matches the operation. Host diagnostics require an
explicitly selected disposable host; they are separate from offline source and
documentation checks. Native qualification runners retain their own frozen
candidate, original-process exit and restoration requirements.

## Common commands

| Operation | Entry point |
| --- | --- |
| Install, remove or verify the source-build registration | [Install-VBAi.ps1](Install-VBAi.ps1), [Uninstall-VBAi.ps1](Uninstall-VBAi.ps1), [Test-VBAiInstallation.ps1](Test-VBAiInstallation.ps1) |
| Send an explicit bridge request to a selected host | [Invoke-VBAi.ps1](Invoke-VBAi.ps1) |
| Build editor assets, icons or native renderer | [Build-MonacoAssets.ps1](Build-MonacoAssets.ps1), [Build-WindowIcons.ps1](Build-WindowIcons.ps1), [Build-NativeRenderer.ps1](build/Build-NativeRenderer.ps1) |
| Check documentation | [check_docs.py](docs/check_docs.py) |
| Export a collected coverage inventory | [Export-CoverageInventory.ps1](coverage/Export-CoverageInventory.ps1) |
| Run a prepared test script on an isolated desktop | [Invoke-IsolatedTests.ps1](tests/Invoke-IsolatedTests.ps1) |
| Inspect native VBE windows | [Inspect-VbeNativeWindows.ps1](probes/Inspect-VbeNativeWindows.ps1), selecting `Tree`, `Windows` or `Children` |
| Inspect current VBAi controls and VBE menus | [Inspect-VbeUi.ps1](probes/Inspect-VbeUi.ps1), selecting `Chat`, `Settings`, `ComboBoxes` or `Menu` |
| Survey control duplication | [Test-FormControlDuplication.ps1](probes/Test-FormControlDuplication.ps1), selecting the control type |
| Survey frame duplication | [Test-FrameDuplication.ps1](probes/Test-FrameDuplication.ps1), selecting `Empty`, `Labels`, `SimpleChildren` or `Profiled` |
| Capture a standalone utility window | [Render-UtilityWindow.ps1](tests/Render-UtilityWindow.ps1), selecting `About`, `CrashReport` or `Update` |

Examples and ownership requirements for diagnostic probes are in
[the probe guide](probes/README.md). Build/test commands and native opt-ins are
maintained in [the testing guide](../tests/README.md).

The utility renderer opens and activates a standalone window on the interactive
desktop and writes PNG/JSON results. Run it when that visible interaction is
appropriate. The native-window inventory reads metadata without changing the UI.

## Directory responsibilities

| Directory | Contents |
| --- | --- |
| `build`, `docs`, `coverage`, `localization` | Source-build and offline maintenance commands. |
| `tests` | Managed/UI diagnostic commands and explicit qualification runners. |
| `probes` | Addressed VBE inspection and disposable native experiments. |
| `testing-explorer` | Test Explorer candidate and qualification support. |
| `VbeController`, `XmlDocumentationAudit` | Dedicated controller and XML-documentation utilities. |

Shared commands replace equivalent scripts; scenario-specific assertions stay
with their scenario. Qualification helpers with different ownership, desktop or
exit contracts remain distinct. A missing caller alone does not establish that
a native diagnostic is obsolete.

Machine-local results, frozen scripts, screenshots and binary manifests belong
under `artifacts/`. Preserve those proof sets and their hashes when archiving
historical builds. They are not part of the maintained command catalog.

Obsolete chat/settings scripts that selected windows by the former CodexVBE
titles have been removed. Current inspection uses WinForms automation IDs and
explicit process ownership; no replacement chat-send, approval-change or
window-close action is implied by the read-only inspector.
