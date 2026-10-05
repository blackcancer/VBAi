# Diagnostic probes

## Native windows

`Inspect-VbeNativeWindows.ps1 -HostProcessId <pid> -View Tree` reads the
selected host's VBE root and descendants. `-View Windows` lists its top-level
windows; `-View Children -WindowHandle <handle>` lists descendants of an explicit
window whose process ownership is checked. All views return JSON rows containing
`Handle`, `Parent`, `Class`, `Title` and `Visible`.

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

## Control duplication

`Test-FormControlDuplication.ps1` runs one of seven Excel UserForm control
duplication surveys against an explicitly selected, owned disposable workbook.
It creates a new form and leaves that form in the workbook for inspection.
It does not execute a macro or close the host.

```powershell
powershell.exe -NoProfile -File tools/probes/Test-FormControlDuplication.ps1 `
    -ControlType Label -HostProcessId <owned-excel-pid> `
    -Project VBAProject -Form CodexLabelCopySurvey -OwnedDisposableWorkbook
```

`ControlType` accepts `Label`, `TextBox`, `ComboBox`, `CheckBox`,
`OptionButton`, `ToggleButton`, or `CommandButton`. If `-Form` is omitted, the
probe uses the former control-specific default name. The selected form must
not already exist. `-OwnedDisposableWorkbook` is an explicit assertion by the
operator; the probe checks the Excel PID and fresh form name before mutation.
Each control retains its original seed, bridge command order, revision checks,
and control-specific copy or refusal checks.

## Frame duplication

`Test-FrameDuplication.ps1` groups the empty, labels-only, Label/TextBox and
profiled frame surveys. Select a `-Scenario` of `Empty` (the default), `Labels`,
`SimpleChildren` or `Profiled`. Like the control survey, it requires an explicit
Excel PID, a fresh form name and `-OwnedDisposableWorkbook`; it leaves the new
form available for inspection.

```powershell
powershell.exe -NoProfile -File tools/probes/Test-FrameDuplication.ps1 `
    -Scenario Profiled -HostProcessId <owned-excel-pid> `
    -Project VBAProject -OwnedDisposableWorkbook
```

Each scenario preserves its copy-plan, partial-copy count and unchanged-tree
refusal checks. The profiled case retains all six child types and its unsupported
SpinButton refusal. This replaces the separate label, simple-child and profiled
frame scripts, while retaining the original empty-frame entry point.

Other scripts in this directory are independent diagnostic probes. They may
have different host ownership or recovery requirements; this command does not
dispatch to them.
