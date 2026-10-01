#requires -Version 5.1
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$CandidateAssemblyPath,
    [Parameter(Mandatory = $true)][Guid]$ExpectedMvid,
    [Parameter(Mandatory = $true)][string]$OutputDirectory,
    [switch]$ExecuteOwnedFixture,
    [switch]$ExecutionAndUiOnly,
    [switch]$TestSubsystemOnly
)

# This script never installs/registers the add-in or changes Office trust settings.
# Execution is explicit, uses a new owned Excel PID, and preserves evidence on failure.
$ErrorActionPreference = 'Stop'
if ($ExecutionAndUiOnly -and $TestSubsystemOnly) { throw 'ExecutionAndUiOnly and TestSubsystemOnly are mutually exclusive. No host was created.' }
if ($PSVersionTable.PSEdition -ne 'Desktop') { throw 'Use Windows PowerShell 5.1 (powershell.exe -STA), not PowerShell Core.' }
if (-not [Environment]::Is64BitProcess) { throw 'Use Windows PowerShell x64.' }
if ([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') { throw 'Use powershell.exe -STA.' }
$candidate = (Resolve-Path -LiteralPath $CandidateAssemblyPath).Path
$candidateAssembly = [Reflection.Assembly]::ReflectionOnlyLoadFrom($candidate)
if ($candidateAssembly.GetName().Name -cne 'VBAi' -or $candidateAssembly.ManifestModule.ModuleVersionId -ne $ExpectedMvid) {
    throw 'Candidate identity/MVID mismatch. No host was created.'
}
$registry = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser, [Microsoft.Win32.RegistryView]::Registry64)
try {
    $key = $registry.OpenSubKey('Software\Classes\CLSID\{8E854243-087F-4D6C-9E0E-8622B0E50883}\InprocServer32')
    if ($null -eq $key) { throw 'VBAi x64 per-user registration is missing. No host was created.' }
    try { $codeBase = [string]$key.GetValue('CodeBase') } finally { $key.Dispose() }
    if (-not $codeBase) { throw 'Registered CodeBase is missing. No host was created.' }
    $registered = ([Uri]$codeBase).LocalPath
    # Assembly LoadFrom may unify equal assembly identities, even for different files.
    # Compare bytes before activation, then verify the actual loaded MVID over the pipe.
    if ((Get-FileHash -LiteralPath $registered -Algorithm SHA256).Hash -cne (Get-FileHash -LiteralPath $candidate -Algorithm SHA256).Hash) {
        throw 'Registered DLL bytes differ from the candidate. Review deployment separately; no host was created.'
    }
    $key = $registry.OpenSubKey('Software\Microsoft\VBA\VBE\6.0\Addins64\VBAi.AddIn')
    if ($null -eq $key) { throw 'VBE add-in discovery registration is missing.' }
    try { if ([int]$key.GetValue('LoadBehavior') -ne 3) { throw 'Expected LoadBehavior=3; the script will not change it.' } } finally { $key.Dispose() }
    $key = $registry.OpenSubKey('Software\Classes\CLSID\{5AF2F40B-939B-4CC6-A06C-F0C79841C031}\InprocServer32')
    if ($null -eq $key) { throw 'VBAi.TestRuntime per-user x64 registration is missing; no host was created.' }
    try {
        if ([string]$key.GetValue('Class') -cne 'VBAi.VbaTestRuntime') { throw 'Unexpected test-runtime COM class; no host was created.' }
        $runtimePath = ([Uri][string]$key.GetValue('CodeBase')).LocalPath
        if ((Get-FileHash -LiteralPath $runtimePath -Algorithm SHA256).Hash -cne (Get-FileHash -LiteralPath $candidate -Algorithm SHA256).Hash) {
            throw 'Registered test-runtime DLL differs from the candidate; no host was created.'
        }
    } finally { $key.Dispose() }
} finally { $registry.Dispose() }
if (-not $ExecuteOwnedFixture) {
    [pscustomobject]@{ Candidate = $candidate; Registered = $registered; Mvid = $ExpectedMvid.ToString('D'); Execution = 'NOT_STARTED'; ExecutionAndUiOnly = [bool]$ExecutionAndUiOnly; TestSubsystemOnly = [bool]$TestSubsystemOnly; Next = 'Run with -ExecuteOwnedFixture only after approving activation of this candidate.' }
    return
}

Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes, UIAutomationClientsideProviders, WindowsBase, System.Drawing
# PowerShell can expose HWNDs as opaque panes until the framework's standard proxies are loaded.
# This registration is local to this diagnostic client; it does not change COM or Windows registry settings.
$uiaProviders = [Reflection.Assembly]::Load('UIAutomationClientsideProviders, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35')
Add-Type -ReferencedAssemblies @(
    [System.Windows.Automation.AutomationElement].Assembly.Location,
    [System.Windows.Automation.AutomationProperty].Assembly.Location,
    [System.Windows.Rect].Assembly.Location,
    [System.Drawing.Bitmap].Assembly.Location,
    [System.Diagnostics.Process].Assembly.Location
) -TypeDefinition @'
using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Windows.Automation;
public static class VBAiExplorerWindowProbe {
    [StructLayout(LayoutKind.Sequential)] public struct Rect { public int Left, Top, Right, Bottom; }
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint pid);
    [DllImport("user32.dll")] [return:MarshalAs(UnmanagedType.Bool)] public static extern bool IsWindow(IntPtr hwnd);
    [DllImport("user32.dll")] [return:MarshalAs(UnmanagedType.Bool)] public static extern bool GetWindowRect(IntPtr hwnd, out Rect rect);
    [DllImport("user32.dll")] [return:MarshalAs(UnmanagedType.Bool)] private static extern bool PrintWindow(IntPtr hwnd, IntPtr dc, uint flags);
    public static bool CaptureWindow(IntPtr hwnd, uint expectedPid, string path) {
        uint pid;
        Rect rect;
        GetWindowThreadProcessId(hwnd, out pid);
        if (!IsWindow(hwnd) || pid != expectedPid || !GetWindowRect(hwnd, out rect)) return false;
        int width = rect.Right - rect.Left, height = rect.Bottom - rect.Top;
        if (width <= 0 || height <= 0 || width > 10000 || height > 10000) return false;
        using (var bitmap = new Bitmap(width, height)) {
            using (var graphics = Graphics.FromImage(bitmap)) {
                IntPtr dc = graphics.GetHdc();
                try { if (!PrintWindow(hwnd, dc, 0)) return false; }
                finally { graphics.ReleaseHdc(dc); }
            }
            // A successful API return can still produce an empty host surface.
            int first = bitmap.GetPixel(0, 0).ToArgb();
            bool varied = false;
            for (int y = 0; y < height && !varied; y += Math.Max(1, height / 50))
                for (int x = 0; x < width && !varied; x += Math.Max(1, width / 50))
                    varied = bitmap.GetPixel(x, y).ToArgb() != first;
            if (!varied) return false;
            bitmap.Save(path, ImageFormat.Png);
            return true;
        }
    }
}
public sealed class VBAiExplorerButtonMonitor : IDisposable {
    private readonly AutomationElement button;
    private readonly AutomationPropertyChangedEventHandler handler;
    private readonly object gate = new object();
    private bool disabled, enabledAfterDisabled;
    public bool SawDisabled { get { lock (gate) return disabled; } }
    public bool SawEnabledAfterDisabled { get { lock (gate) return enabledAfterDisabled; } }
    public VBAiExplorerButtonMonitor(AutomationElement button) {
        this.button = button;
        handler = (sender, args) => {
            if (args.Property == AutomationElement.IsEnabledProperty && args.NewValue is bool)
                Observe((bool)args.NewValue);
        };
        // A native managed callback needs no PowerShell runspace on the UIA worker.
        Automation.AddAutomationPropertyChangedEventHandler(button, TreeScope.Element,
            handler, AutomationElement.IsEnabledProperty);
    }
    public void Observe(bool enabled) {
        lock (gate) {
            if (!enabled) disabled = true;
            else if (disabled) enabledAfterDisabled = true;
        }
    }
    public void Dispose() { Automation.RemoveAutomationPropertyChangedEventHandler(button, handler); }
}
'@
$root = Join-Path ([IO.Path]::GetFullPath($OutputDirectory)) ([Guid]::NewGuid().ToString('N'))
[IO.Directory]::CreateDirectory($root) | Out-Null
$bridge = Join-Path (Split-Path -Parent $PSScriptRoot) 'Invoke-VBAi.ps1'
$report = [ordered]@{ Candidate = $candidate; Registered = $registered; ExpectedMvid = $ExpectedMvid.ToString('D'); Scope = 'Registered in-process Excel bridge and exact HWND/PID native explorer actions; LLM guards are not exercised by this bridge'; StartedUtc = [DateTime]::UtcNow.ToString('o'); UiActions = 'NOT_TESTED'; LlmGuards = 'NOT_TESTED'; CopyReport = 'NOT_TESTED' }
$report.ExecutionAndUiOnly = [bool]$ExecutionAndUiOnly
$report.TestSubsystemOnly = [bool]$TestSubsystemOnly
if ($ExecutionAndUiOnly) { $report.Scope = 'Execution and UI only: registered in-process Excel bridge and exact HWND/PID native explorer actions. BeforeSave and measured coverage are NOT_TESTED_BY_SCOPE. LLM guards are not exercised by this bridge.' }
elseif ($TestSubsystemOnly) { $report.Scope = 'VBA testing subsystem qualification: registered in-process Excel bridge, measured procedure coverage, human/JSON reports and exact HWND/PID native explorer actions. The general BeforeSave positive control is NOT_TESTED_BY_SCOPE; retained event counters and probe results are observations only and do not prove handler suppression. LLM guards are not exercised by this bridge.' }
else { $report.Scope = 'Full qualification: registered in-process Excel bridge, BeforeSave, measured procedure coverage and exact HWND/PID native explorer actions. LLM guards are not exercised by this bridge.' }
$report.ScriptSha256 = (Get-FileHash -LiteralPath $PSCommandPath -Algorithm SHA256).Hash
$report.CandidateSha256 = (Get-FileHash -LiteralPath $candidate -Algorithm SHA256).Hash
$existing = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
$application = $null; $workbook = $null; $workbooks = $null; $ownedSheets = $null; $ownedSheet = $null; $ownedCells = $null; $beforeSaveCell = $null; $vbe = $null; $vbeWindow = $null; $project = $null; $components = $null; $module = $null; $code = $null; $production = $null; $productionCode = $null; $workbookModule = $null; $workbookCode = $null; $ownedProcess = $null; $owned = $false
$nativeExecutionUnsettled = $false
$coverageId = $null
$beforeSaveBaseline = $null
$scenarioError = $null; $cleanupErrors = @()
function Call-Bridge([hashtable]$Request, [switch]$AllowFailure) {
    $reply = & $bridge -HostProcessId $script:ownedProcess.Id -RequestJson ($Request | ConvertTo-Json -Depth 15 -Compress) -ResponseTimeoutSeconds 30 | ConvertFrom-Json
    if (-not $AllowFailure -and -not $reply.Ok) { throw "$($Request.Command): $($reply.Error)" }
    if ($AllowFailure) { return $reply }
    return $reply.Data
}
function Require([bool]$Condition, [string]$Message) { if (-not $Condition) { throw $Message } }
function Save-Evidence([string]$Name, $Value) { $Value | ConvertTo-Json -Depth 30 | Set-Content -LiteralPath (Join-Path $root $Name) -Encoding UTF8 }

function Require-OwnedFixtureReady {
    $script:ownedProcess.Refresh()
    Require (-not $script:ownedProcess.HasExited) 'Owned Excel lifetime ended; fixture inspection refused.'
    [uint32]$fixturePid = 0
    [void][VBAiExplorerWindowProbe]::GetWindowThreadProcessId([IntPtr][long]$script:application.Hwnd, [ref]$fixturePid)
    Require ($fixturePid -eq $script:ownedProcess.Id -and $script:workbooks.Count -eq 1) 'Owned Excel PID/document set changed; fixture inspection refused.'
    Require ([IO.Path]::GetFullPath([string]$script:workbook.FullName) -ieq [IO.Path]::GetFullPath($script:path)) 'Owned workbook path changed; fixture inspection refused.'
    $currentProject = $script:workbook.VBProject
    $expectedIdentity = [IntPtr]::Zero; $currentIdentity = [IntPtr]::Zero
    try {
        $expectedIdentity = [Runtime.InteropServices.Marshal]::GetIUnknownForObject($script:project)
        $currentIdentity = [Runtime.InteropServices.Marshal]::GetIUnknownForObject($currentProject)
        Require ($currentIdentity -eq $expectedIdentity -and [int]$script:project.Mode -eq 2) 'Owned VBA project identity/mode changed; fixture inspection refused.'
    } finally {
        if ($currentIdentity -ne [IntPtr]::Zero) { [void][Runtime.InteropServices.Marshal]::Release($currentIdentity) }
        if ($expectedIdentity -ne [IntPtr]::Zero) { [void][Runtime.InteropServices.Marshal]::Release($expectedIdentity) }
        if ($null -ne $currentProject -and [Runtime.InteropServices.Marshal]::IsComObject($currentProject)) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($currentProject) }
    }
    Require ([string]$script:code.Lines(1, $script:code.CountOfLines) -ceq $script:original -and [string]$script:productionCode.Lines(1, $script:productionCode.CountOfLines) -ceq $script:productionOriginal -and [string]$script:workbookCode.Lines(1, $script:workbookCode.CountOfLines) -ceq $script:workbookOriginal) 'Owned fixture source changed; counter execution refused.'
    $currentCatalog = Call-Bridge @{ Command = 'discover_vba_tests'; Project = $script:path }
    Require ($currentCatalog.ExpectedProjectVersion -ceq $script:catalog.ExpectedProjectVersion) 'Owned project revision, support or references changed; counter execution refused.'
}
function Get-UiNames {
    $localization = Join-Path (Split-Path -Parent (Split-Path -Parent $PSScriptRoot)) 'src/VBAi/Localization'
    $resources = @(Get-ChildItem -LiteralPath $localization -Filter 'UiStrings*.resx' -File)
    Require ($resources.Count -eq 13) 'Expected the complete set of 13 native UI resource dictionaries.'
    $names = @{}
    foreach ($key in @('Run visible scope', 'Readable report', 'LLM report (JSON)', 'Human-readable test report', 'Compact test report for LLM', 'Test run completed.', 'Passed', 'Failed', 'Error')) {
        $values = @()
        foreach ($resource in $resources) {
            [xml]$xml = Get-Content -LiteralPath $resource.FullName -Raw -Encoding UTF8
            $entry = @($xml.root.data | Where-Object { $_.name -ceq $key })
            Require ($entry.Count -eq 1 -and [string]$entry[0].value) ('Missing native UI resource: ' + $key + ' in ' + $resource.Name)
            $values += [string]$entry[0].value
        }
        $names[$key] = @($values | Select-Object -Unique)
    }
    return $names
}

function Get-OwnedExplorerElement {
    [uint32]$pidNow = 0
    [void][VBAiExplorerWindowProbe]::GetWindowThreadProcessId($script:window, [ref]$pidNow)
    Require ([VBAiExplorerWindowProbe]::IsWindow($script:window) -and $pidNow -eq $script:actualPid) 'Explorer HWND ownership changed; native UI action refused.'
    $element = [System.Windows.Automation.AutomationElement]::FromHandle($script:window)
    Require ($null -ne $element -and $element.Current.ProcessId -eq $script:actualPid) 'UIAutomation root does not belong to the exact owned explorer HWND/PID.'
    return $element
}

function Find-ExplorerElement([System.Windows.Automation.ControlType[]]$Type, [string[]]$Names) {
    $element = Get-OwnedExplorerElement
    $conditions = @($Type | ForEach-Object { [System.Windows.Automation.PropertyCondition]::new([System.Windows.Automation.AutomationElement]::ControlTypeProperty, $_) })
    $condition = if ($conditions.Count -eq 1) { $conditions[0] } else { [System.Windows.Automation.OrCondition]::new([System.Windows.Automation.Condition[]]$conditions) }
    $matches = @()
    foreach ($child in $element.FindAll([System.Windows.Automation.TreeScope]::Descendants, $condition)) {
        if ($child.Current.ProcessId -eq $script:actualPid -and $Names -ccontains $child.Current.Name) { $matches += $child }
    }
    Require ($matches.Count -eq 1) ('Expected one exact owned explorer control; found ' + $matches.Count + ': ' + ($Names -join ' / '))
    return $matches[0]
}

function Read-ExplorerReport([string]$TabKey, [string]$ReportKey) {
    $tab = Find-ExplorerElement ([System.Windows.Automation.ControlType]::TabItem) $script:uiNames[$TabKey]
    Require (-not $tab.Current.IsOffscreen -and $tab.Current.IsEnabled) 'Report tab is not visible/enabled in the owned explorer.'
    $selection = $tab.GetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern)
    Require ($null -ne $selection) 'Report tab does not expose SelectionItemPattern.'
    $selection.Select()
    Require ($selection.Current.IsSelected) 'Native report tab selection failed.'
    $box = Find-ExplorerElement @([System.Windows.Automation.ControlType]::Edit, [System.Windows.Automation.ControlType]::Document) $script:uiNames[$ReportKey]
    Require (-not $box.Current.IsOffscreen) 'Report text is not visible after native tab selection.'
    [object]$value = $null
    if ($box.TryGetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern, [ref]$value)) {
        Require ($value.Current.IsReadOnly) 'Native report ValuePattern must be read-only.'
        return [string]$value.Current.Value
    }
    # The standard WinForms provider exposes read-only multiline edits as documents with TextPattern.
    [object]$textPattern = $null
    Require ($box.TryGetCurrentPattern([System.Windows.Automation.TextPattern]::Pattern, [ref]$textPattern)) 'Native report must expose ValuePattern or TextPattern.'
    $range = $textPattern.DocumentRange
    Require ($range.GetAttributeValue([System.Windows.Automation.TextPattern]::IsReadOnlyAttribute) -eq $true) 'Native report TextPattern must be read-only.'
    $content = $range.GetText(1048577)
    Require ($content.Length -le 1048576) 'Native report exceeds the bounded qualification reader.'
    return [string]$content
}

try {
    $application = New-Object -ComObject Excel.Application
    [uint32]$actualPid = 0
    [void][VBAiExplorerWindowProbe]::GetWindowThreadProcessId([IntPtr][long]$application.Hwnd, [ref]$actualPid)
    Require ($actualPid -gt 0 -and $existing -notcontains [int]$actualPid) 'Excel reused an existing PID; no workbook will be created and that process will not be closed.'
    $owned = $true; $ownedProcess = Get-Process -Id $actualPid
    [void]$ownedProcess.get_Handle() # Retain the exact native process for exit-code observation.
    $report.Pid = [int]$actualPid; $report.ExcelVersion = [string]$application.Version
    $application.Visible = $true
    $vbe = $application.VBE
    $vbeWindow = $vbe.MainWindow
    $vbeWindow.Visible = $true
    # status is the only retryable call here; it has no mutation or execution effect.
    $status = $null
    for ($attempt = 0; $attempt -lt 20 -and $null -eq $status; $attempt++) {
        try { $status = Call-Bridge @{ Command = 'status' } } catch { if ($attempt -eq 19) { throw }; Start-Sleep -Milliseconds 250 }
    }
    Require ($status.HostProcessId -eq $actualPid -and $status.ProcessBitness -eq 64 -and [Guid]$status.AssemblyModuleVersionId -eq $ExpectedMvid) 'Loaded add-in PID/bitness/MVID mismatch.'
    Require ([IO.Path]::GetFullPath($status.AssemblyPath) -ieq [IO.Path]::GetFullPath($registered)) 'Loaded add-in path differs from the registered path.'
    $report.Identity = $status; Save-Evidence 'identity.json' $status
    $workbooks = $application.Workbooks
    Require ($workbooks.Count -eq 0) 'Owned Excel opened unexpected workbooks; no synthetic document will be added.'
    $workbook = $workbooks.Add()
    $path = Join-Path $root 'registered-test-explorer.xlsm'
    $workbook.SaveAs($path, 52)
    $ownedSheets = $workbook.Worksheets
    $ownedSheet = $ownedSheets.Item(1)
    $ownedCells = $ownedSheet.Cells
    $beforeSaveCell = $ownedCells.Item(1, 1)
    $beforeSaveCell.Value2 = 0
    $project = $workbook.VBProject # Existing Trust Center policy must permit this; never change it.
    $components = $project.VBComponents
    $module = $components.Add(1); $module.Name = 'RegisteredExplorerTests'
    $source = @"
Option Explicit
'@TestModule
Private cleanupCalls As Long
Private beforeSaveCalls As Long
'@TestInitialize
Public Sub InitializeTest()
End Sub
'@TestCleanup
Public Sub CleanupTest()
    cleanupCalls = cleanupCalls + 1
End Sub
'@TestMethod
Public Function BooleanPass() As Boolean
    BooleanPass = (RegisteredCalculator.AddOne(1) = 2)
End Function
'@TestMethod
Public Function BooleanFail() As Boolean
    BooleanFail = False
End Function
'@TestMethod
Public Sub SwallowedAssertion()
    On Error Resume Next
    VBAiTestSupport.Fail "Swallowed failure must remain failed."
End Sub
'@TestMethod
Public Sub RuntimeError()
    Err.Raise 5, "RegisteredExplorerFixture", "Deliberate runtime error."
End Sub
Public Function ReadCleanupCount() As Long
    ReadCleanupCount = cleanupCalls
End Function
Public Sub RecordBeforeSave()
    beforeSaveCalls = beforeSaveCalls + 1
End Sub
Public Function ReadBeforeSaveCount() As Long
    ReadBeforeSaveCount = beforeSaveCalls
End Function
Public Sub ResetBeforeSaveCount()
    beforeSaveCalls = 0
End Sub
"@
    $code = $module.CodeModule; $code.AddFromString($source)
    $original = [string]$code.Lines(1, $code.CountOfLines)
    $production = $components.Add(1); $production.Name = 'RegisteredCalculator'
    $productionCode = $production.CodeModule
    $productionCode.AddFromString(@'
Option Explicit
Public Function AddOne(ByVal value As Long) As Long
    AddOne = value + 1
End Function
Public Function NeverCalled() As Long
    NeverCalled = 99
End Function
'@)
    $productionOriginal = [string]$productionCode.Lines(1, $productionCode.CountOfLines)
    # Add an event only to the fresh disposable workbook's exact document module.
    $workbookModule = $components.Item([string]$workbook.CodeName)
    Require ([int]$workbookModule.Type -eq 100) 'The owned workbook document module is not a document component.'
    $workbookCode = $workbookModule.CodeModule
    Require ($workbookCode.CountOfLines -eq 0) 'The disposable workbook document module unexpectedly contains source.'
    $workbookCode.AddFromString(@'
Option Explicit
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    Me.Worksheets(1).Cells(1, 1).Value2 = Me.Worksheets(1).Cells(1, 1).Value2 + 1
    RegisteredExplorerTests.RecordBeforeSave
End Sub
'@)
    $workbookOriginal = [string]$workbookCode.Lines(1, $workbookCode.CountOfLines)
    $catalog = Call-Bridge @{ Command = 'discover_vba_tests'; Project = $path }
    Save-Evidence 'discovery-before.json' $catalog
    $preview = Call-Bridge @{ Command = 'preview_vba_test_support'; Project = $path }
    [IO.File]::WriteAllText((Join-Path $root 'reviewed-support.bas'), [string]$preview.Text, [Text.UTF8Encoding]::new($false))
    $installed = Call-Bridge @{ Command = 'install_vba_test_support'; Project = $path; ExpectedProjectVersion = $preview.ExpectedProjectVersion; ExpectedMode = 2; Text = $preview.Text }
    Save-Evidence 'support-install.json' $installed
    $catalog = Call-Bridge @{ Command = 'discover_vba_tests'; Project = $path }
    $ids = @($catalog.Modules | ForEach-Object { $_.Tests } | ForEach-Object { $_.Id })
    Require ($ids.Count -eq 4) 'Expected exactly four synthetic tests.'
    if ($ExecutionAndUiOnly -or $TestSubsystemOnly) {
        $report.BeforeSave = [ordered]@{ Result = 'NOT_TESTED_BY_SCOPE'; HandlerQualifiedByOwnedSave = 'NOT_TESTED_BY_SCOPE'; Observation = 'OBSERVATION_ONLY'; HandlerSuppression = 'NOT_TESTED_BY_SCOPE'; OriginalUnchangedByCoverage = 'NOT_TESTED'; EnableEventsRestored = 'NOT_TESTED' }
        if ($TestSubsystemOnly) { Require-OwnedFixtureReady }
        $workbook.Save() # Save only the owned synthetic fixture; no positive-control claim.
        $report.BeforeSave.PersistentCellAfterFixtureSave = $beforeSaveCell.Value2
    } else {
        $report.BeforeSave = [ordered]@{ Stage = 'WarmRead'; HandlerQualifiedByOwnedSave = $false; OriginalUnchangedByCoverage = 'NOT_TESTED'; EnableEventsRestored = 'NOT_TESTED' }
        $report.BeforeSave.SavedBeforeWarmRead = [bool]$workbook.Saved
        $report.BeforeSave.EnableEventsBeforeWarmRead = [bool]$application.EnableEvents
        $report.BeforeSave.PersistentCell = [string]$ownedSheet.Name + '!A1'
        $report.BeforeSave.PersistentCellBeforeWarmRead = $beforeSaveCell.Value2
        # Warm the VBA getter before the event: a first invocation can compile the fresh project.
        $beforeSaveBeforeControl = [int]$application.Run(("'" + $path.Replace("'", "''") + "'!RegisteredExplorerTests.ReadBeforeSaveCount"))
        $report.BeforeSave.CounterBeforeSave = $beforeSaveBeforeControl
        $report.BeforeSave.PersistentCellAfterWarmRead = $beforeSaveCell.Value2
        $report.BeforeSave.SavedAfterWarmRead = [bool]$workbook.Saved
        $report.BeforeSave.EnableEventsBeforeSave = [bool]$application.EnableEvents
        Save-Evidence 'before-save-control.json' $report.BeforeSave
        Require ($report.BeforeSave.EnableEventsBeforeSave) 'Owned Excel events must initially be enabled to qualify the BeforeSave handler.'
        Require ($beforeSaveBeforeControl -eq 0) 'The fresh synthetic BeforeSave counter is not zero before the positive control.'
        # Saved is a read/write Boolean. Mark only this owned fixture dirty so Save has work.
        # https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.saved
        $report.BeforeSave.Stage = 'MarkOwnedFixtureDirty'
        $workbook.Saved = $false
        $report.BeforeSave.SavedImmediatelyBeforeSave = [bool]$workbook.Saved
        $report.BeforeSave.PersistentCellImmediatelyBeforeSave = $beforeSaveCell.Value2
        $report.BeforeSave.WorkbookCodeNameBeforeSave = [string]$workbook.CodeName
        $report.BeforeSave.DocumentModuleNameBeforeSave = [string]$workbookModule.Name
        $report.BeforeSave.ProjectModeBeforeSave = [int]$project.Mode
        $documentLinesBeforeSave = [int]$workbookCode.CountOfLines
        $documentSourceBeforeSave = if ($documentLinesBeforeSave -gt 0) { [string]$workbookCode.Lines(1, $documentLinesBeforeSave) } else { '' }
        $report.BeforeSave.DocumentSourceMatchesBeforeSave = $documentSourceBeforeSave -ceq $workbookOriginal
        if (-not $report.BeforeSave.DocumentSourceMatchesBeforeSave) {
            [IO.File]::WriteAllText((Join-Path $root 'before-save-document-expected.bas'), $workbookOriginal, [Text.UTF8Encoding]::new($false))
            [IO.File]::WriteAllText((Join-Path $root 'before-save-document-before.bas'), $documentSourceBeforeSave, [Text.UTF8Encoding]::new($false))
            $report.BeforeSave.DocumentSourceBeforeMismatchFile = 'before-save-document-before.bas'
        }
        Save-Evidence 'before-save-control.json' $report.BeforeSave
        Require (-not $report.BeforeSave.SavedImmediatelyBeforeSave) 'The owned fixture could not be marked unsaved for its positive BeforeSave control.'
        $report.BeforeSave.Stage = 'ExplicitOwnedSave'
        $report.BeforeSave.SaveStartedUtc = [DateTime]::UtcNow.ToString('o')
        $workbook.Save() # Persist only the owned synthetic fixture for reproducibility.
        $report.BeforeSave.SaveReturnedUtc = [DateTime]::UtcNow.ToString('o')
        $report.BeforeSave.SavedImmediatelyAfterSave = [bool]$workbook.Saved
        $report.BeforeSave.EnableEventsImmediatelyAfterSave = [bool]$application.EnableEvents
        # Read the exact owned cell before executing any getter that could reset VBA globals.
        $report.BeforeSave.PersistentCellAfterSaveBeforeGetter = $beforeSaveCell.Value2
        $report.BeforeSave.PersistentCounterDelta = [double]$report.BeforeSave.PersistentCellAfterSaveBeforeGetter - [double]$report.BeforeSave.PersistentCellImmediatelyBeforeSave
        $report.BeforeSave.WorkbookCodeNameAfterSave = [string]$workbook.CodeName
        $report.BeforeSave.DocumentModuleNameAfterSave = [string]$workbookModule.Name
        $report.BeforeSave.ProjectModeAfterSaveBeforeGetter = [int]$project.Mode
        $documentLinesAfterSave = [int]$workbookCode.CountOfLines
        $documentSourceAfterSave = if ($documentLinesAfterSave -gt 0) { [string]$workbookCode.Lines(1, $documentLinesAfterSave) } else { '' }
        $report.BeforeSave.DocumentSourceMatchesAfterSaveBeforeGetter = $documentSourceAfterSave -ceq $workbookOriginal
        if (-not $report.BeforeSave.DocumentSourceMatchesAfterSaveBeforeGetter) {
            [IO.File]::WriteAllText((Join-Path $root 'before-save-document-expected.bas'), $workbookOriginal, [Text.UTF8Encoding]::new($false))
            [IO.File]::WriteAllText((Join-Path $root 'before-save-document-after.bas'), $documentSourceAfterSave, [Text.UTF8Encoding]::new($false))
            $report.BeforeSave.DocumentSourceAfterMismatchFile = 'before-save-document-after.bas'
        }
        Save-Evidence 'before-save-control.json' $report.BeforeSave
        $beforeSaveObserved = [int]$application.Run(("'" + $path.Replace("'", "''") + "'!RegisteredExplorerTests.ReadBeforeSaveCount"))
        $report.BeforeSave.CounterAfterSave = $beforeSaveObserved
        $report.BeforeSave.PositiveControlCounterDelta = $beforeSaveObserved - $beforeSaveBeforeControl
        $report.BeforeSave.SavedAfterCounterRead = [bool]$workbook.Saved
        $report.BeforeSave.EnableEventsAfterCounterRead = [bool]$application.EnableEvents
        $report.BeforeSave.PersistentCellAfterGetter = $beforeSaveCell.Value2
        Save-Evidence 'before-save-control.json' $report.BeforeSave
        Require ($report.BeforeSave.SavedImmediatelyAfterSave) 'The positive-control Save returned without a saved owned workbook.'
        Require ($report.BeforeSave.EnableEventsAfterCounterRead -eq $report.BeforeSave.EnableEventsBeforeSave) 'The positive-control Save changed Excel EnableEvents.'
        Require ($beforeSaveObserved -eq 1) ('The synthetic BeforeSave handler did not execute exactly once for the explicit owned save; observed counter=' + $beforeSaveObserved + '.')
        $report.BeforeSave.ExplicitSaveCount = $beforeSaveObserved
        $report.BeforeSave.HandlerQualifiedByOwnedSave = $true
        $report.BeforeSave.Stage = 'ResetOwnedCounter'
        [void]$application.Run(("'" + $path.Replace("'", "''") + "'!RegisteredExplorerTests.ResetBeforeSaveCount"))
        $beforeSaveBaseline = [int]$application.Run(("'" + $path.Replace("'", "''") + "'!RegisteredExplorerTests.ReadBeforeSaveCount"))
        $report.BeforeSave.Baseline = $beforeSaveBaseline
        Save-Evidence 'before-save-control.json' $report.BeforeSave
        Require ($beforeSaveBaseline -eq 0) 'The synthetic BeforeSave counter did not reset for the isolated coverage observation.'
        $report.BeforeSave.Stage = 'CoverageBaselineReady'
    }
    $nativeExecutionUnsettled = $true
    $started = Call-Bridge @{ Command = 'run_vba_tests'; Project = $path; ExpectedProjectVersion = $catalog.ExpectedProjectVersion; ExpectedMode = 2; Items = $ids }
    $runId = $started.Query
    $deadline = [DateTime]::UtcNow.AddSeconds(45)
    do {
        $state = Call-Bridge @{ Command = 'vba_test_run_status'; Project = $path; Query = $runId; Action = 'compact' }
        if (-not $state.Pending) { break }
        Start-Sleep -Milliseconds 100
    } while ([DateTime]::UtcNow -lt $deadline)
    Save-Evidence 'results-compact.json' $state
    if (-not $state.Pending -and $null -ne $state.Report) { $nativeExecutionUnsettled = $state.Report.uncertain -ne $false }
    Require (-not $state.Pending) 'Native run did not complete before the observation deadline. No retry/reset was performed.'
    Require ($state.State -eq 'Completed') 'Registered native run did not complete normally.'
    Require (-not $state.Report.uncertain -and -not $state.Report.error -and $state.Report.tests.Count -eq 4) 'Run report lacks four verified results.'
    $expectedResults = @(
        @{ Name = 'BooleanPass'; Outcome = 'Passed'; Error = 0 },
        @{ Name = 'BooleanFail'; Outcome = 'Failed'; Error = 0 },
        @{ Name = 'SwallowedAssertion'; Outcome = 'Failed' },
        @{ Name = 'RuntimeError'; Outcome = 'Error'; Error = 5 }
    )
    foreach ($expected in $expectedResults) {
        $actual = @($state.Report.tests | Where-Object { $_.name -ceq ('RegisteredExplorerTests.' + $expected.Name) })
        Require ($actual.Count -eq 1 -and $actual[0].outcome -ceq $expected.Outcome) ('Incorrect native verdict: ' + $expected.Name)
        if ($expected.ContainsKey('Error')) { Require ($actual[0].error -eq $expected.Error) ('Incorrect native error number: ' + $expected.Name) }
    }
    $human = Call-Bridge @{ Command = 'vba_test_run_status'; Project = $path; Query = $runId; Action = 'human' }
    [IO.File]::WriteAllText((Join-Path $root 'results-human.txt'), [string]$human.Report, [Text.UTF8Encoding]::new($false))
    Require ([string]$code.Lines(1, $code.CountOfLines) -ceq $original) 'Test source changed during execution.'
    if ($TestSubsystemOnly) { Require-OwnedFixtureReady }
    $cleanupCount = [int]$application.Run(("'" + $path.Replace("'", "''") + "'!RegisteredExplorerTests.ReadCleanupCount"))
    Require ($cleanupCount -eq 4) 'Independent cleanup count differs from four completed tests.'
    $report.CleanupCalls = $cleanupCount
    if ($ExecutionAndUiOnly) {
        $report.Coverage = @{ Result = 'NOT_TESTED_BY_SCOPE'; Available = $false }
        $report.MeasuredCoverage = @{ Result = 'NOT_TESTED_BY_SCOPE'; Available = $false }
    } else {
        $coverage = Call-Bridge @{ Command = 'vba_test_coverage'; Project = $path }
        Require ($coverage.Project -ceq $path -and $coverage.ExpectedProjectVersion -ceq $catalog.ExpectedProjectVersion) 'Coverage preview does not identify the exact owned source path/revision.'
        Require ($coverage.Supported -and $coverage.Metric -ceq 'Procedure' -and $coverage.DenominatorKnown -and $coverage.EligibleProcedureCount -eq 3) 'Coverage preview must expose AddOne, NeverCalled and Workbook_BeforeSave as three production procedures.'
        Require ($coverage.ProbeTotal -eq 3 -and $coverage.Probes.Count -eq 3 -and $null -eq $coverage.NextOffset) 'Coverage preview does not contain all three procedure mappings.'
        $mappingExpected = @(
            @{ Module = 'RegisteredCalculator'; Procedure = 'AddOne'; Kind = 'Function'; Line = [int]$productionCode.ProcBodyLine('AddOne', 0) },
            @{ Module = 'RegisteredCalculator'; Procedure = 'NeverCalled'; Kind = 'Function'; Line = [int]$productionCode.ProcBodyLine('NeverCalled', 0) },
            @{ Module = [string]$workbookModule.Name; Procedure = 'Workbook_BeforeSave'; Kind = 'Sub'; Line = [int]$workbookCode.ProcBodyLine('Workbook_BeforeSave', 0) }
        )
        foreach ($mapping in $mappingExpected) {
            $probe = @($coverage.Probes | Where-Object { $_.Module -ceq $mapping.Module -and $_.Procedure -ceq $mapping.Procedure })
            Require ($probe.Count -eq 1 -and $probe[0].Kind -ceq $mapping.Kind -and $probe[0].OriginalLine -eq $mapping.Line -and $probe[0].OriginalColumn -gt 0 -and $probe[0].Id) ('Invalid original-source coverage mapping: ' + $mapping.Module + '.' + $mapping.Procedure)
        }
        Require (@($coverage.Probes | ForEach-Object { $_.Id } | Select-Object -Unique).Count -eq 3) 'Coverage preview probe identities are not unique.'
        Save-Evidence 'coverage-preview.json' $coverage
        $report.Coverage = $coverage
        $eventsBeforeCoverage = [bool]$application.EnableEvents
        if ($TestSubsystemOnly) { Require-OwnedFixtureReady }
        $beforeSaveBeforeCoverage = [int]$application.Run(("'" + $path.Replace("'", "''") + "'!RegisteredExplorerTests.ReadBeforeSaveCount"))
        if ($TestSubsystemOnly) {
            # Baseline is read only after the completed test batch; there is no warm-up,
            # positive event control or counter reset in the subsystem-only scenario.
            $beforeSaveBaseline = $beforeSaveBeforeCoverage
            $report.BeforeSave.Baseline = $beforeSaveBaseline
            $report.BeforeSave.PersistentCellBeforeCoverage = $beforeSaveCell.Value2
        }
        Require (($TestSubsystemOnly -or $eventsBeforeCoverage) -and $beforeSaveBeforeCoverage -eq $beforeSaveBaseline) 'Original BeforeSave counter or Excel event state changed before coverage.'
        $nativeExecutionUnsettled = $true
        $coverageStarted = Call-Bridge @{ Command = 'run_vba_tests'; Project = $path; ExpectedProjectVersion = $catalog.ExpectedProjectVersion; ExpectedMode = 2; Items = $ids; Action = 'coverage' }
        $coverageId = $coverageStarted.Query
        $deadline = [DateTime]::UtcNow.AddSeconds(60)
        do {
            $coverageState = Call-Bridge @{ Command = 'vba_test_run_status'; Project = $path; Query = $coverageId; Action = 'compact' }
            if (-not $coverageState.Pending) { break }
            Start-Sleep -Milliseconds 100
        } while ([DateTime]::UtcNow -lt $deadline)
        Save-Evidence 'coverage-results-compact.json' $coverageState
        if (-not $coverageState.Pending -and $null -ne $coverageState.Report) { $nativeExecutionUnsettled = $coverageState.Report.uncertain -ne $false }
        Require (-not $coverageState.Pending -and $coverageState.State -ceq 'Completed') 'Coverage clone did not complete normally; no retry was performed.'
        $measured = Call-Bridge @{ Command = 'vba_test_coverage'; Project = $path; Query = $coverageId }
        Save-Evidence 'coverage-results-query.json' $measured
        Require ($measured.Report.Original -ceq $path -and $measured.Report.Revision -ceq $catalog.ExpectedProjectVersion -and $measured.Report.Metric -ceq 'Procedure') 'Measured coverage does not identify the exact owned original path/revision/metric.'
        Require ($measured.Available -and $measured.Report.Available -and $measured.Report.Complete -and $measured.Report.DenominatorKnown) 'Measured procedure coverage is incomplete.'
        Require ($measured.Report.Eligible -eq 3 -and $measured.Report.Hit -eq 1 -and [Math]::Abs([double]$measured.Report.Percent - (100.0 / 3.0)) -lt 0.000001) 'Expected measured procedure coverage 1/3; BeforeSave and NeverCalled must not be entered.'
        Require ([string]$productionCode.Lines(1, $productionCode.CountOfLines) -ceq $productionOriginal -and [string]$code.Lines(1, $code.CountOfLines) -ceq $original -and [string]$workbookCode.Lines(1, $workbookCode.CountOfLines) -ceq $workbookOriginal) 'Coverage changed original source.'
        $eventsAfterCoverage = [bool]$application.EnableEvents
        if ($TestSubsystemOnly) { Require-OwnedFixtureReady }
        $beforeSaveAfterCoverage = [int]$application.Run(("'" + $path.Replace("'", "''") + "'!RegisteredExplorerTests.ReadBeforeSaveCount"))
        Require ($eventsAfterCoverage -eq $eventsBeforeCoverage) 'Coverage did not restore the exact original Excel EnableEvents state.'
        Require ($beforeSaveAfterCoverage -eq $beforeSaveBeforeCoverage) 'The original synthetic BeforeSave counter changed during coverage.'
        $beforeSaveProbe = @($measured.Report.Hits | Where-Object { $_.Probe.Module -ceq [string]$workbookModule.Name -and $_.Probe.Procedure -ceq 'Workbook_BeforeSave' })
        Require ($beforeSaveProbe.Count -eq 1 -and -not $beforeSaveProbe[0].Entered) 'The coverage copy executed Workbook_BeforeSave unexpectedly.'
        foreach ($previewProbe in $coverage.Probes) {
            $hitMapping = @($measured.Report.Hits | Where-Object { $_.Probe.Id -ceq $previewProbe.Id })
            Require ($hitMapping.Count -eq 1 -and $hitMapping[0].Probe.Module -ceq $previewProbe.Module -and $hitMapping[0].Probe.Procedure -ceq $previewProbe.Procedure -and $hitMapping[0].Probe.OriginalLine -eq $previewProbe.OriginalLine) 'Measured coverage mapping differs from the reviewed original-source preview.'
        }
        $report.BeforeSave.BeforeCoverage = $beforeSaveBeforeCoverage
        $report.BeforeSave.AfterCoverage = $beforeSaveAfterCoverage
        $report.BeforeSave.PersistentCellAfterCoverage = $beforeSaveCell.Value2
        $report.BeforeSave.EnableEventsBefore = $eventsBeforeCoverage
        $report.BeforeSave.EnableEventsAfter = $eventsAfterCoverage
        $report.BeforeSave.CoverageBeforeSaveEntered = [bool]$beforeSaveProbe[0].Entered
        $report.BeforeSave.OriginalUnchangedByCoverage = 'PASS'
        $report.BeforeSave.EnableEventsRestored = 'PASS'
        if ($TestSubsystemOnly) { Require-OwnedFixtureReady }
        Require ([int]$application.Run(("'" + $path.Replace("'", "''") + "'!RegisteredExplorerTests.ReadCleanupCount")) -eq 4) 'Coverage executed tests in the original workbook instead of the clone.'
        $report.MeasuredCoverage = $measured
    }
    $opened = Call-Bridge @{ Command = 'show_vba_test_explorer'; Project = $path }
    Save-Evidence 'explorer-window.json' $opened
    if ($opened.Hwnd) {
        $window = [IntPtr][long]$opened.Hwnd; [uint32]$windowPid = 0
        [void][VBAiExplorerWindowProbe]::GetWindowThreadProcessId($window, [ref]$windowPid)
        $bounds = New-Object VBAiExplorerWindowProbe+Rect
        Require ([VBAiExplorerWindowProbe]::IsWindow($window) -and $windowPid -eq $actualPid) 'Explorer HWND does not belong to the owned Excel PID.'
        Require ([VBAiExplorerWindowProbe]::GetWindowRect($window, [ref]$bounds) -and $bounds.Right -gt $bounds.Left -and $bounds.Bottom -gt $bounds.Top) 'Explorer window bounds are invalid.'
        $report.ExplorerWindow = @{ Hwnd = [long]$window; Pid = $windowPid; Docked = $opened.Docked; Bounds = $bounds }
    } else { throw 'show_vba_test_explorer must return Hwnd for independent native-window qualification.' }

    # Showing the same explorer twice must reuse the exact owned native window.
    $shownAgain = Call-Bridge @{ Command = 'show_vba_test_explorer'; Project = $path }
    Save-Evidence 'explorer-window-reused.json' $shownAgain
    Require ([long]$shownAgain.Hwnd -eq [long]$window) 'Second show created a different explorer HWND.'
    [void](Get-OwnedExplorerElement)
    # Initialize UIA against the exact owned HWND before installing the standard client proxies.
    [System.Windows.Automation.ClientSettings]::RegisterClientSideProviderAssembly($uiaProviders.GetName())
    $report.ExplorerWindow.Reused = $true
    $uiNames = Get-UiNames
    Save-Evidence 'native-ui-resource-names.json' $uiNames
    $uiInventory = @()
    foreach ($child in (Get-OwnedExplorerElement).FindAll([System.Windows.Automation.TreeScope]::Descendants, [System.Windows.Automation.Condition]::TrueCondition)) {
        Require ($child.Current.ProcessId -eq $actualPid) 'Foreign UIAutomation descendant; explorer action refused.'
        $uiInventory += [pscustomobject]@{ Name = $child.Current.Name; Type = $child.Current.ControlType.ProgrammaticName;
            Hwnd = $child.Current.NativeWindowHandle; AutomationId = $child.Current.AutomationId;
            ClassName = $child.Current.ClassName; IsOffscreen = $child.Current.IsOffscreen; IsEnabled = $child.Current.IsEnabled; Patterns = @($child.GetSupportedPatterns() | ForEach-Object ProgrammaticName) }
        if ($uiInventory.Count -ge 256) { break }
    }
    Save-Evidence 'native-ui-descendants.json' $uiInventory
    $runButton = Find-ExplorerElement ([System.Windows.Automation.ControlType]::Button) $uiNames['Run visible scope']
    Require (-not $runButton.Current.IsOffscreen -and $runButton.Current.IsEnabled) 'Run visible scope must be visible and enabled in the exact owned explorer.'
    $invoke = $runButton.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern)
    Require ($null -ne $invoke) 'Run visible scope does not expose InvokePattern.'
    $report.NativeUi = [ordered]@{ Hwnd = [long]$window; Pid = [int]$actualPid; Button = $runButton.Current.Name; Invocations = 0; ObservationLimitSeconds = 60 }
    $monitor = New-Object VBAiExplorerButtonMonitor -ArgumentList $runButton
    try {
        $deadline = [DateTime]::UtcNow.AddSeconds(60)
        [void](Get-OwnedExplorerElement)
        $report.NativeUi.Invocations = 1
        $report.NativeUi.StartedUtc = [DateTime]::UtcNow.ToString('o')
        # Never retry this invocation, including when the native result is still pending.
        $nativeExecutionUnsettled = $true
        $invoke.Invoke()
        do {
            [void](Get-OwnedExplorerElement)
            $monitor.Observe($runButton.Current.IsEnabled)
            if ($monitor.SawDisabled -and $monitor.SawEnabledAfterDisabled -and $runButton.Current.IsEnabled) { break }
            Start-Sleep -Milliseconds 50
        } while ([DateTime]::UtcNow -lt $deadline)
        Require ([DateTime]::UtcNow -le $deadline -and $monitor.SawDisabled -and $monitor.SawEnabledAfterDisabled -and $runButton.Current.IsEnabled) 'Native UI run did not disable/re-enable within 60 seconds. No macro was retried.'
    } finally {
        $report.NativeUi.SawDisabled = $monitor.SawDisabled
        $report.NativeUi.SawEnabledAfterDisabled = $monitor.SawEnabledAfterDisabled
        $report.NativeUi.ObservationFinishedUtc = [DateTime]::UtcNow.ToString('o')
        $monitor.Dispose()
    }
    $completedLabel = Find-ExplorerElement ([System.Windows.Automation.ControlType]::Text) $uiNames['Test run completed.']
    Require (-not $completedLabel.Current.IsOffscreen) 'Native explorer does not display its completed status.'
    $report.NativeUi.CompletedStatus = $completedLabel.Current.Name
    $uiHuman = Read-ExplorerReport 'Readable report' 'Human-readable test report'
    $uiJson = Read-ExplorerReport 'LLM report (JSON)' 'Compact test report for LLM'
    [IO.File]::WriteAllText((Join-Path $root 'native-ui-results-human.txt'), $uiHuman, [Text.UTF8Encoding]::new($false))
    [IO.File]::WriteAllText((Join-Path $root 'native-ui-results-compact.json'), $uiJson, [Text.UTF8Encoding]::new($false))
    $uiCompact = $uiJson | ConvertFrom-Json
    Require ($uiCompact.v -eq 1 -and $uiCompact.run -and $uiCompact.run -cne $runId -and $uiCompact.run -cne $coverageId) 'Native UI report does not identify a new run.'
    Require ($uiCompact.revision -ceq $catalog.ExpectedProjectVersion -and -not $uiCompact.uncertain -and -not $uiCompact.error -and $uiCompact.tests.Count -eq 4) 'Native UI report lacks four verified results for the reviewed revision.'
    Require ($null -ne $uiCompact.coverage -and $uiCompact.coverage.available -eq $false -and $uiCompact.coverage.reason -ceq 'No coverage run was requested.') 'Normal native UI execution incorrectly claims measured coverage.'
    $idDifferences = @(Compare-Object -ReferenceObject $ids -DifferenceObject @($uiCompact.tests | ForEach-Object { $_.id }) -CaseSensitive)
    Require ($idDifferences.Count -eq 0) 'Native UI report test IDs differ from the discovered selection.'
    Require ($uiHuman.Contains([string]$uiCompact.run) -and $uiHuman.Contains([string]$catalog.ExpectedProjectVersion)) 'Human report identifies a different run/revision from the compact report.'
    foreach ($expected in $expectedResults) {
        $name = 'RegisteredExplorerTests.' + $expected.Name
        $actual = @($uiCompact.tests | Where-Object { $_.name -ceq $name })
        $descriptor = @($catalog.Modules | ForEach-Object { $_.Tests } | Where-Object { $_.Module -ceq 'RegisteredExplorerTests' -and $_.Procedure -ceq $expected.Name })
        Require ($actual.Count -eq 1 -and $descriptor.Count -eq 1 -and $actual[0].id -ceq $descriptor[0].Id -and $actual[0].outcome -ceq $expected.Outcome) ('Incorrect native UI verdict/identity: ' + $expected.Name)
        if ($expected.ContainsKey('Error')) { Require ($actual[0].error -eq $expected.Error) ('Incorrect native UI error number: ' + $expected.Name) }
        $symbol = if ($expected.Outcome -ceq 'Passed') { [char]0x2713 } else { [char]0x2717 }
        $lines = @($uiHuman -split '\r?\n' | Where-Object { $_.Contains(' | ' + $name + ' | ') })
        Require ($lines.Count -eq 1) ('Human UI report lacks exactly one result: ' + $expected.Name)
        $localizedVerdicts = @($uiNames[$expected.Outcome] | Where-Object { $lines[0].StartsWith([string]$symbol + ' ' + $_ + ' | ', [StringComparison]::Ordinal) })
        Require ($localizedVerdicts.Count -ge 1) ('Human UI report outcome/symbol differs from compact JSON: ' + $expected.Name)
    }
    # Read-only bridge corroboration cannot initiate or repeat the UI batch.
    $uiState = Call-Bridge @{ Command = 'vba_test_run_status'; Project = $path; Query = $uiCompact.run; Action = 'compact' }
    Save-Evidence 'native-ui-run-status.json' $uiState
    if (-not $uiState.Pending -and $null -ne $uiState.Report) { $nativeExecutionUnsettled = $uiState.Report.uncertain -ne $false }
    Require (-not $uiState.Pending -and $uiState.State -ceq 'Completed' -and $uiState.Report.run -ceq $uiCompact.run -and -not $uiState.Report.uncertain) 'Native UI batch is not independently verified as completed.'
    Require ([string]$productionCode.Lines(1, $productionCode.CountOfLines) -ceq $productionOriginal -and [string]$code.Lines(1, $code.CountOfLines) -ceq $original -and [string]$workbookCode.Lines(1, $workbookCode.CountOfLines) -ceq $workbookOriginal) 'Native UI execution changed fixture source.'
    # Only inspect the independent VBA counter after the UI batch has completed.
    if ($TestSubsystemOnly) { Require-OwnedFixtureReady }
    $cleanupCount = [int]$application.Run(("'" + $path.Replace("'", "''") + "'!RegisteredExplorerTests.ReadCleanupCount"))
    Require ($cleanupCount -eq 8) 'Independent cleanup count must be eight after bridge and native UI batches.'
    $report.NativeUi.RunId = $uiCompact.run
    $report.NativeUi.CleanupCalls = $cleanupCount
    $report.NativeUi.Reports = @('native-ui-results-human.txt', 'native-ui-results-compact.json')
    $report.CleanupCalls = $cleanupCount
    $report.UiActions = 'PASS'
    # PrintWindow captures this exact HWND only; unavailable rendering is explicit.
    $report.Screenshot = 'NOT_AVAILABLE'
    try {
        [void](Read-ExplorerReport 'Readable report' 'Human-readable test report')
        $screenshot = Join-Path $root 'native-ui-explorer.png'
        if ([VBAiExplorerWindowProbe]::CaptureWindow($window, $actualPid, $screenshot)) {
            $report.Screenshot = @{ File = 'native-ui-explorer.png'; Method = 'PrintWindow'; Hwnd = [long]$window; Pid = [int]$actualPid }
        }
    } catch { $report.ScreenshotReason = $_.Exception.Message }
    $code.AddFromString("' changed after native run")
    $nativeExecutionUnsettled = $true
    $refused = Call-Bridge @{ Command = 'run_vba_tests'; Project = $path; ExpectedProjectVersion = $catalog.ExpectedProjectVersion; ExpectedMode = 2; Items = $ids } -AllowFailure
    Require (-not $refused.Ok) 'Stale native batch was accepted.'
    $nativeExecutionUnsettled = $false
    $report.StaleRefusal = $refused.Error
    $report.Result = 'PASS_WITH_EXPLICIT_SCOPE'
} catch {
    $scenarioError = $_; $report.Result = 'FAILED'; $report.Error = $_.Exception.ToString()
} finally {
    $cleanupUnsettled = $nativeExecutionUnsettled
    $report.Cleanup = [ordered]@{ NativeExecutionUnsettled = $nativeExecutionUnsettled; CloseAttempted = $false; CloseReturned = $false; QuitAttempted = $false; QuitReturned = $false }
    # Close only after rechecking the COM application's current HWND ownership.
    if ($owned -and $null -ne $application -and -not $nativeExecutionUnsettled) {
        try {
            [uint32]$closePid = 0
            [void][VBAiExplorerWindowProbe]::GetWindowThreadProcessId([IntPtr][long]$application.Hwnd, [ref]$closePid)
            Require ($closePid -eq $ownedProcess.Id) 'Excel ownership changed; automatic close refused.'
            if ($null -ne $workbook) {
                $report.Cleanup.CloseAttempted = $true
                $workbook.Close($false)
                $report.Cleanup.CloseReturned = $true
            }
            $report.Cleanup.RemainingWorkbookCount = [int]$workbooks.Count
            Require ($report.Cleanup.RemainingWorkbookCount -eq 0) 'Unexpected documents remain; automatic Quit refused.'
            $report.Cleanup.QuitAttempted = $true
            $report.Cleanup.QuitStartedUtc = [DateTime]::UtcNow.ToString('o')
            $application.Quit()
            $report.Cleanup.QuitReturned = $true
            $report.Cleanup.QuitReturnedUtc = [DateTime]::UtcNow.ToString('o')
        } catch {
            $cleanupErrors += $_.Exception.Message
            $cleanupUnsettled = $true
            $report.Cleanup.CleanupOutcomeUnsettled = $true
        }
    } elseif ($owned -and $nativeExecutionUnsettled) {
        $cleanupErrors += 'Native execution remains pending or uncertain; Close/Quit were not attempted. Inspect the exact owned PID before recovery.'
    }
    if ($owned -and $cleanupUnsettled) {
        # Keep this client alive: interpreter exit would also drop its RCWs and process handle.
        # Observe only the captured process handle; never reopen a PID or retry a native mutation.
        $report.Result = 'FAILED'
        $report.NormalExit = $false
        $report.WaitForExitReturned = $null
        $report.HasExited = $null
        $report.ExitCode = $null
        $report.Cleanup.RetainedOwnership = $true
        $report.Cleanup.OwnedProcessId = $ownedProcess.Id
        $report.Cleanup.ComReferencesRetained = $true
        $report.Cleanup.ProcessHandleRetained = $true
        $report.Cleanup.ReleasedOwnedComReferences = 0
        $report.Cleanup.RetentionStartedUtc = [DateTime]::UtcNow.ToString('o')
        $cleanupErrors += 'Native outcome is unsettled. Ownership is retained until the exact owned process exits externally; no Close/Quit retry is permitted.'
        $retentionMarker = [ordered]@{ RetainedOwnership = $true; OwnedProcessId = $ownedProcess.Id; NativeExecutionUnsettled = $nativeExecutionUnsettled; CleanupOutcomeUnsettled = $cleanupUnsettled; StartedUtc = $report.Cleanup.RetentionStartedUtc }
        $report.CleanupErrors = $cleanupErrors
        try { Save-Evidence 'qualification.json' $report }
        catch { $cleanupErrors += ('Retention qualification could not be persisted: ' + $_.Exception.Message) }
        try { Save-Evidence 'retained-ownership.json' $retentionMarker }
        catch { $cleanupErrors += ('Retention marker could not be persisted: ' + $_.Exception.Message) }
        Write-Output ("Ownership retained for exact owned PID {0}. This client will remain alive until that process exits. Evidence: {1}" -f $ownedProcess.Id, $root)
        $ownedExitObserved = $false
        $retentionObservationError = $null
        while (-not $ownedExitObserved) {
            try {
                if ($ownedProcess.get_HasExited()) {
                    $report.HasExited = $true
                    $report.ExitCode = $ownedProcess.get_ExitCode()
                    $report.Cleanup.RetainedProcessExitObservedUtc = [DateTime]::UtcNow.ToString('o')
                    $retentionMarker.ExitObservedUtc = $report.Cleanup.RetainedProcessExitObservedUtc
                    $retentionMarker.ExitCode = $report.ExitCode
                    $ownedExitObserved = $true
                }
            } catch {
                # Observation failure cannot authorize dropping ownership or ending this client.
                if ($retentionObservationError -cne $_.Exception.Message) {
                    $retentionObservationError = $_.Exception.Message
                    $cleanupErrors += ('Retained owned process exit could not be observed: ' + $retentionObservationError)
                    $report.Cleanup.ProcessObservationError = $retentionObservationError
                    $report.CleanupErrors = $cleanupErrors
                    try { Save-Evidence 'qualification.json' $report } catch { }
                }
            }
            if (-not $ownedExitObserved) { Start-Sleep -Seconds 1 }
        }
        # Exit permits releasing references to the dead process, but never resolves the unknown run/cleanup.
        $report.Cleanup.RetainedOwnershipReleasedAfterExit = $true
        try { Save-Evidence 'retained-ownership.json' $retentionMarker }
        catch { $cleanupErrors += ('Final retention evidence could not be persisted: ' + $_.Exception.Message) }
    }
    $releasedReferences = 0
    foreach ($item in @($workbookCode, $workbookModule, $productionCode, $production, $code, $module, $components, $project, $beforeSaveCell, $ownedCells, $ownedSheet, $ownedSheets, $workbook, $workbooks, $vbeWindow, $vbe, $application)) {
        if ($null -ne $item -and [Runtime.InteropServices.Marshal]::IsComObject($item)) {
            try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($item); $releasedReferences++ } catch { $cleanupErrors += $_.Exception.Message }
        }
    }
    $item = $null
    $report.Cleanup.ReleasedOwnedComReferences = $releasedReferences
    $workbookCode = $null; $workbookModule = $null; $productionCode = $null; $production = $null; $code = $null; $module = $null; $components = $null; $project = $null; $beforeSaveCell = $null; $ownedCells = $null; $ownedSheet = $null; $ownedSheets = $null; $workbook = $null; $workbooks = $null; $vbeWindow = $null; $vbe = $null; $application = $null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect()
    if ($ownedProcess) {
        if ($owned -and $cleanupUnsettled) {
            $report.Cleanup.ComReferencesRetained = $false
            $ownedProcess.Dispose()
            $report.Cleanup.ProcessHandleRetained = $false
            # NormalExit intentionally remains false even when the externally observed exit code is zero.
        } else {
            $report.WaitForExitReturned = $null
            $report.HasExited = $null
            $report.ExitCode = $null
            $report.NormalExit = $false
            try {
                $report.Cleanup.WaitForExitStartedUtc = [DateTime]::UtcNow.ToString('o')
                $exited = $ownedProcess.WaitForExit(15000)
                $ownedProcess.Refresh()
                $report.WaitForExitReturned = $exited
                $report.HasExited = $ownedProcess.get_HasExited()
                $report.ExitCode = if ($report.HasExited) { $ownedProcess.get_ExitCode() } else { $null }
                $report.Cleanup.WaitForExitFinishedUtc = [DateTime]::UtcNow.ToString('o')
                $report.NormalExit = $report.HasExited -and $report.ExitCode -eq 0
                if (-not $report.HasExited) { $cleanupErrors += 'Owned Excel is still running after the observation deadline; it was not killed.' }
                elseif ($report.ExitCode -ne 0) { $cleanupErrors += ('Owned Excel exited with code ' + $report.ExitCode + '; it was not killed.') }
            } catch {
                $report.Cleanup.ProcessObservationError = $_.Exception.Message
                $cleanupErrors += ('Owned Excel exit status could not be verified: ' + $_.Exception.Message)
            } finally { $ownedProcess.Dispose() }
        }
    }
    $report.CleanupErrors = $cleanupErrors
    $report.FinishedUtc = [DateTime]::UtcNow.ToString('o')
    Save-Evidence 'qualification.json' $report
    Write-Output "Qualification evidence: $root"
}
if ($scenarioError) { throw $scenarioError }
if ($cleanupErrors.Count) { throw ($cleanupErrors -join '; ') }
