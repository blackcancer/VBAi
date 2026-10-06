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
| Inspect code, debugger panes, MSAA, references or Object Browser | [Inspect-VbeAccessibility.ps1](probes/Inspect-VbeAccessibility.ps1), selecting `-Target` |
| UserForm controls, properties, rollback and duplication | [Test-UserFormControls.ps1](probes/Test-UserFormControls.ps1), selecting `-Scenario`; duplication adds `-ControlType` or `-FrameScenario` |
| UserForm list items and initialization code | [Test-UserFormLists.ps1](probes/Test-UserFormLists.ps1), selecting `Append`, `ReadPages`, `Remove` or `InitializeWithoutEvent` |
| UserForm lifecycle, selection, clipboard and history | [Test-UserFormNative.ps1](probes/Test-UserFormNative.ps1), selecting `-Scenario` |
| UserForm COM inspection and disposable control surveys | [Inspect-UserForm.ps1](probes/Inspect-UserForm.ps1), selecting `-Scenario` |
| Debugger, breakpoints, execution and watches | [Test-Debugger.ps1](probes/Test-Debugger.ps1), selecting `-Scenario`; `BreakpointStage` adds `-Stage` |
| Code panes, round trips, definitions and history | [Test-CodeEditor.ps1](probes/Test-CodeEditor.ps1), selecting `-Scenario` |
| VBE events and functional extensions | [Test-VbeEvents.ps1](probes/Test-VbeEvents.ps1), selecting `-Scenario` |
| Excel code, persistence and bookmarks | [Test-ExcelCode.ps1](tests/Test-ExcelCode.ps1), selecting `-Scenario` |
| Selected preloaded SOLIDWORKS macro diagnostics | [Test-SolidWorks.ps1](tests/Test-SolidWorks.ps1), selecting `-Scenario` |

Examples and ownership requirements for diagnostic probes are in
[the probe guide](probes/README.md). Build/test commands and native opt-ins are
maintained in [the testing guide](../tests/README.md).

The domain commands contain their scenario implementations and share equivalent
setup and transport. They do not retain the old scripts behind forwarding
wrappers. Select one scenario per invocation; there is no implicit native `All`.
[VbeProbe.Common.ps1](probes/VbeProbe.Common.ps1) shares bridge transport and
identical Excel setup while each scenario retains its own cleanup and oracles.

## UI checks and rendering

| Execution boundary | Command | Scenarios |
| --- | --- | --- |
| Host-free workflow, Designer construction or metadata | [Test-UiOffline.ps1](tests/Test-UiOffline.ps1) | `ChatWorkflow`, `ChatDesigner`, `WinFormsDesigners`, `ProjectMetadata` |
| Standalone UI that may show windows or initialize WebView | [Test-Ui.ps1](tests/Test-Ui.ps1) | `ChatUx`, `EditorAppearance`, `FormRecoveryCard`, `MultilingualWindows`, `UiReview`, `DisplayProfiles` |
| Registered control, disposable Excel or selected Visual Studio | [Test-UiHost.ps1](tests/Test-UiHost.ps1) | `ChatControlActivation`, `ExcelFormLayouts`, `VisualStudioWinForms` |
| Explicit standalone rendering | [Render-Ui.ps1](tests/Render-Ui.ps1) | `ChatDesignerViews`, `ChatUx`, `CompactUi`, `UtilityWindow` |

[UiProbe.psm1](tests/UiProbe.psm1) shares reflection, assertions, bitmap capture and
resource disposal. Importing it opens no assembly, window or host. Rendering
requires a suitable test desktop. `UtilityWindow` opens and activates a window
and additionally selects `-Window About`, `CrashReport` or `Update`. The
native-window inventory reads metadata without changing the UI.

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

Generated build outputs and caches also belong under `artifacts/`. The
XML-documentation audit now defaults to `artifacts/build/XmlDocumentationAudit`
and `artifacts/obj/XmlDocumentationAudit`; an explicit `BuildOutputRoot` still
selects its build output. The old Git/provider smoke sources are maintained in
[tests/](../tests/README.md). Their abandoned outputs under `tools/tests/Git` and
`tools/tests/Providers` are not active projects.

Obsolete chat/settings scripts that selected windows by the former CodexVBE
titles have been removed. Current inspection uses WinForms automation IDs and
explicit process ownership; no replacement chat-send, approval-change or
window-close action is implied by the read-only inspector.
