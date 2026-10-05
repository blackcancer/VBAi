# Diagnostic probes

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

Other scripts in this directory are independent diagnostic probes. They may
have different host ownership or recovery requirements; this command does not
dispatch to them.
