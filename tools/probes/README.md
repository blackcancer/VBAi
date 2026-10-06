# Diagnostic probes

## Native windows

`Inspect-VbeNativeWindows.ps1 -HostProcessId <pid> -View Tree` reads the
selected host's VBE root and descendants. `-View Windows` lists its top-level
windows; `-View Children -WindowHandle <handle>` lists descendants of an explicit
window whose process ownership is checked. All views return JSON rows containing
`Handle`, `Parent`, `Class`, `Title` and `Visible`.

`-View Panes -PaneNames <names>` retains the pane/child text inventory while
sharing native enumeration and ownership checks with the other views.

This replaces `Inspect-VbeWindowTree.ps1`, `Inspect-VbeTopWindows.ps1` and
`Inspect-VbeDialogChildren.ps1`. It does not open, activate or close a window.

## VBAi controls and menus

`Inspect-VbeUi.ps1 -Area Chat -HostProcessId <pid>` inspects the unique
`ChatWindow` descendant of the selected host's main window. `-Area Settings`
uses `LlmSettingsWindow`. Selection uses automation IDs and process ownership,
so it does not depend on the localized window title. If the window is owned
rather than a descendant, supply its exact `-WindowHandle`; that root is then
validated and inspected directly.

`-Area ComboBoxes` requires the explicit chat window handle and lists at most
100 items per combo, refusing text longer than 4,096 characters. `-Area Menu`
reports matching menu controls and their available patterns without invoking
or expanding them. These views read the selected UI without activating it.

These modes replace the separate chat, settings, combo and menu UI Automation
inspectors. Native MSAA and debugger inspectors remain separate because their
targets and optional actions differ.

## Accessibility, code and debugger inspection

`Inspect-VbeAccessibility.ps1 -Target <target> -HostProcessId <pid>` shares
VBE-root discovery and type loading for `CodePane`, `CodePaneMsaa`,
`CodeSelection`, `DebugTree`, `Locals`, `MsaaChildren`, `Menu`, `References`
and `ObjectBrowser`. Code pane targets accept `-CodeWindowName`; locals/MSAA
accept `-PaneName`. `Menu` and `References` require the explicit owned
`-WindowHandle`. Reference text is capped at 4,096 characters per item.

Reading does not activate a control. Object Browser selection requires
`-Class`, `-Member` and `-AcknowledgeUiAction`; `-BrowserName` can select its
localized title. Menu `-OpenMenu`/`-InvokeItem` and debugger `-Invoke` or
`-ClickNativeButton` also require `-AcknowledgeUiAction`. `InvokeItem` requires
`OpenMenu`. A failed or uncertain activation ends the command without trying
another matching control.

## UserForm scenarios

| Command | Scenarios |
| --- | --- |
| `Test-UserFormControls.ps1` | `AddControlRollback`, `AddNestedRollback`, `ComboColumnGuard`, `ControlEnumChoices`, `PropertyStatus`, `SingleNodeSetter`, `SpinProperties`, `SpinWriteGuard`, `FrameCopyPlan`, `FrameDuplication`, `ControlDuplication` |
| `Test-UserFormLists.ps1` | `Append`, `ReadPages`, `Remove`, `InitializeWithoutEvent` |
| `Test-UserFormNative.ps1` | `Selection`, `Lifecycle`, `Clipboard`, `ClipboardRecoveryDiscovery`, `History` |
| `Inspect-UserForm.ps1` | `Explore`, `Nested`, `ExportProperties`, `SurveyControls`, `FormProperties` |

Controls and list mutations require an explicit Excel PID and
`-OwnedDisposableWorkbook`. Scenarios retain their fresh/empty form, design-mode,
revision and refusal checks. `PropertyStatus` and `SpinWriteGuard` also require
the expected COM CodeBase. `SingleNodeSetter` requires `ControlType`, `Property`
and `ValueJson`. Native scenarios own their disposable Excel fixture and retain
their individual trust, clipboard, save/reopen and exit contracts. Inspection
`Explore -Stage Create` requires `-OwnedDisposableWorkbook`; `SurveyControls`
retains its `-AcknowledgeDisposableProject` gate.

### Control duplication

`Test-UserFormControls.ps1 -Scenario ControlDuplication` runs one of seven Excel UserForm control
duplication surveys against an explicitly selected, owned disposable workbook.
It creates a new form and leaves that form in the workbook for inspection.
It does not execute a macro or close the host.

```powershell
powershell.exe -NoProfile -STA -File tools/probes/Test-UserFormControls.ps1 `
    -Scenario ControlDuplication -ControlType Label -HostProcessId <owned-excel-pid> `
    -Project VBAProject -Form CodexLabelCopySurvey -OwnedDisposableWorkbook
```

`ControlType` accepts `Label`, `TextBox`, `ComboBox`, `CheckBox`,
`OptionButton`, `ToggleButton`, or `CommandButton`. If `-Form` is omitted, the
probe uses the former control-specific default name. The selected form must
not already exist. `-OwnedDisposableWorkbook` is an explicit assertion by the
operator; the probe checks the Excel PID and fresh form name before mutation.
Each control retains its original seed, bridge command order, revision checks,
and control-specific copy or refusal checks.

### Frame duplication

`Test-UserFormControls.ps1 -Scenario FrameDuplication` groups the empty, labels-only,
Label/TextBox and profiled frame surveys. Select a `-FrameScenario` of `Empty` (the default), `Labels`,
`SimpleChildren` or `Profiled`. Like the control survey, it requires an explicit
Excel PID, a fresh form name and `-OwnedDisposableWorkbook`; it leaves the new
form available for inspection.

```powershell
powershell.exe -NoProfile -STA -File tools/probes/Test-UserFormControls.ps1 `
    -Scenario FrameDuplication -FrameScenario Profiled -HostProcessId <owned-excel-pid> `
    -Project VBAProject -OwnedDisposableWorkbook
```

Each scenario preserves its copy-plan, partial-copy count and unchanged-tree
refusal checks. The profiled case retains all six child types and its unsupported
SpinButton refusal. This replaces the separate label, simple-child and profiled
frame scripts.

## Code, events and debugger

`Test-Debugger.ps1`, `Test-CodeEditor.ps1` and `Test-VbeEvents.ps1` group
debugger, editor and event scenarios. PID mutations require
`-OwnedDisposableHost`. Assembly scenarios take explicit `-AssemblyPath` and
`-OutputDirectory`; registered variants retain `-UseBridge`. Each invocation
runs one scenario with its original fixture lifetime and cleanup.

Use `../tests/Test-ExcelCode.ps1` for Excel persistence, code and navigation.
Use `../tests/Test-SolidWorks.ps1` for a selected, already-open disposable
SOLIDWORKS macro; it does not launch or close the host. The registered Excel
dispatcher remains `../tests/Test-RegisteredExcelNavigation.ps1` and selects
the corresponding domain command and explicit scenario.

Other scripts in this directory are independent diagnostic probes. They may
have different host ownership or recovery requirements; this command does not
dispatch to them.
