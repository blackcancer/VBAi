#requires -Version 5.1
[CmdletBinding(DefaultParameterSetName = 'Preview')]
param(
    [Parameter(Mandatory = $true, ParameterSetName = 'Preview')]
    [Parameter(Mandatory = $true, ParameterSetName = 'Apply')][string]$CandidateAssemblyPath,
    [Parameter(Mandatory = $true, ParameterSetName = 'Preview')]
    [Parameter(Mandatory = $true, ParameterSetName = 'Apply')][Guid]$ExpectedMvid,
    [Parameter(Mandatory = $true, ParameterSetName = 'Apply')][switch]$Apply,
    [Parameter(Mandatory = $true, ParameterSetName = 'Restore')][switch]$Restore,
    [Parameter(Mandatory = $true, ParameterSetName = 'Restore')][string]$BackupPath,
    [Parameter(ParameterSetName = 'Apply')][string]$ReportPath
)

# Changes future COM activation only. Never installs files, changes Office trust,
# starts/stops a host, migrates data, or automatically restores registration.
$ErrorActionPreference = 'Stop'
if ($PSVersionTable.PSEdition -ne 'Desktop') { throw 'Use Windows PowerShell 5.1.' }
if (-not [Environment]::Is64BitProcess) { throw 'Use x64 PowerShell.' }
$registry = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser, [Microsoft.Win32.RegistryView]::Registry64)
$addIn = 'Software\Classes\CLSID\{8E854243-087F-4D6C-9E0E-8622B0E50883}'
$chat = 'Software\Classes\CLSID\{0F4D723B-97D8-42E5-9B31-70646B97C8D2}'
$runtime = 'Software\Classes\CLSID\{5AF2F40B-939B-4CC6-A06C-F0C79841C031}'
$runtimeProg = 'Software\Classes\VBAi.TestRuntime'
$roots = @($addIn, $chat, $runtime, $runtimeProg)

function Sorted-Names($Values) {
    [string[]]$ordered = @($Values)
    [Array]::Sort($ordered, [StringComparer]::Ordinal)
    return ,$ordered
}
function Read-Tree([string]$Path) {
    $key = $registry.OpenSubKey($Path)
    if ($null -eq $key) { return [pscustomobject]@{ Path = $Path; Exists = $false; Values = @(); Children = @() } }
    try {
        $values = @()
        foreach ($name in (Sorted-Names $key.GetValueNames())) {
            $kind = $key.GetValueKind($name)
            $data = $key.GetValue($name, $null, [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames)
            if ($kind -eq [Microsoft.Win32.RegistryValueKind]::Binary -or $kind -eq [Microsoft.Win32.RegistryValueKind]::None) { $data = [Convert]::ToBase64String([byte[]]$data) }
            elseif ($kind -eq [Microsoft.Win32.RegistryValueKind]::MultiString) { $data = @([string[]]$data) }
            $values += [pscustomobject]@{ Name = $name; Kind = $kind.ToString(); Data = $data }
        }
        $children = @()
        foreach ($name in (Sorted-Names $key.GetSubKeyNames())) { $children += Read-Tree ($Path + '\' + $name) }
        return [pscustomobject]@{ Path = $Path; Exists = $true; Values = $values; Children = $children }
    } finally { $key.Dispose() }
}
function Read-All { return @($roots | ForEach-Object { Read-Tree $_ }) }
function Fingerprint($Trees) {
    $bytes = [Text.Encoding]::UTF8.GetBytes(($Trees | ConvertTo-Json -Depth 100 -Compress))
    $hash = [Security.Cryptography.SHA256]::Create()
    try { return [Convert]::ToBase64String($hash.ComputeHash($bytes)) } finally { $hash.Dispose() }
}
function Write-Tree($Tree) {
    if (-not $Tree.Exists) { return }
    $key = $registry.CreateSubKey([string]$Tree.Path)
    try {
        foreach ($value in $Tree.Values) {
            $kind = [Microsoft.Win32.RegistryValueKind][Enum]::Parse([Microsoft.Win32.RegistryValueKind], [string]$value.Kind)
            $data = $value.Data
            if ($kind -eq [Microsoft.Win32.RegistryValueKind]::Binary -or $kind -eq [Microsoft.Win32.RegistryValueKind]::None) { $data = [Convert]::FromBase64String([string]$data) }
            elseif ($kind -eq [Microsoft.Win32.RegistryValueKind]::MultiString) { $data = [string[]]@($data) }
            elseif ($kind -eq [Microsoft.Win32.RegistryValueKind]::DWord) { $data = [int]$data }
            elseif ($kind -eq [Microsoft.Win32.RegistryValueKind]::QWord) { $data = [long]$data }
            else { $data = [string]$data }
            $key.SetValue([string]$value.Name, $data, $kind)
        }
    } finally { $key.Dispose() }
    foreach ($child in $Tree.Children) { Write-Tree $child }
}
function Require-TreeScope($Tree, [string]$ExpectedPath) {
    if ([string]$Tree.Path -cne $ExpectedPath) { throw 'Backup contains an out-of-scope registry path.' }
    foreach ($child in $Tree.Children) {
        $prefix = $ExpectedPath + '\'
        if (-not ([string]$child.Path).StartsWith($prefix, [StringComparison]::Ordinal)) { throw 'Backup child path is outside its parent.' }
        $name = ([string]$child.Path).Substring($prefix.Length)
        if (-not $name -or $name.Contains('\')) { throw 'Backup child path must be one immediate registry child.' }
        Require-TreeScope $child ($prefix + $name)
    }
}
function Flatten-Trees($Trees) {
    $flat = @{}
    function Add-Node($Node) {
        $flat['K|' + $Node.Path] = [string][bool]$Node.Exists
        foreach ($value in $Node.Values) { $flat['V|' + $Node.Path + [char]0 + $value.Name] = ([pscustomobject]@{ Kind = $value.Kind; Data = $value.Data } | ConvertTo-Json -Depth 100 -Compress) }
        foreach ($child in $Node.Children) { Add-Node $child }
    }
    foreach ($tree in $Trees) { Add-Node $tree }
    return $flat
}
function Owned-Difference($Before, $After, $AllowedKeys, $AllowedValues) {
    $old = Flatten-Trees $Before; $new = Flatten-Trees $After
    foreach ($name in $old.Keys) {
        if (-not $new.ContainsKey($name)) { return $false }
        if ($old[$name] -ceq $new[$name]) { continue }
        if ($name.StartsWith('K|')) { if (-not $AllowedKeys.ContainsKey($name) -or $new[$name] -cne 'True') { return $false } }
        elseif (-not $AllowedValues.ContainsKey($name) -or $new[$name] -cne $AllowedValues[$name]) { return $false }
    }
    foreach ($name in $new.Keys) {
        if ($old.ContainsKey($name)) { continue }
        if ($name.StartsWith('K|')) { if (-not $AllowedKeys.ContainsKey($name) -or $new[$name] -cne 'True') { return $false } }
        elseif (-not $AllowedValues.ContainsKey($name) -or $new[$name] -cne $AllowedValues[$name]) { return $false }
    }
    return $true
}
function Read-Value([string]$Path, [string]$Name) {
    $key = $registry.OpenSubKey($Path)
    if ($null -eq $key) { return $null }
    try { return $key.GetValue($Name, $null, [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames) }
    finally { $key.Dispose() }
}
function Require-Server([string]$Root, [string]$Class, [string]$ProgId, [string]$Clsid) {
    if ((Read-Value ('Software\Classes\' + $ProgId + '\CLSID') '') -ine $Clsid) { throw "Missing or foreign HKCU ProgID: $ProgId" }
    if ((Read-Value ($Root + '\ProgId') '') -cne $ProgId) { throw "Missing or foreign reverse ProgID: $ProgId" }
    $server = $Root + '\InprocServer32'
    $key = $registry.OpenSubKey($server)
    if ($null -eq $key) { throw "Missing HKCU server: $ProgId" }
    try {
        $paths = @($server)
        foreach ($name in $key.GetSubKeyNames()) {
            $version = $null
            if ([Version]::TryParse($name, [ref]$version)) { $paths += $server + '\' + $name }
        }
    } finally { $key.Dispose() }
    foreach ($path in $paths) {
        if ((Read-Value $path 'Class') -cne $Class -or -not ([string](Read-Value $path 'Assembly')).StartsWith('VBAi,', [StringComparison]::Ordinal) -or (Read-Value $path '') -ine 'mscoree.dll') {
            # Version keys do not have an unnamed mscoree value in standard RegAsm output.
            if ($path -eq $server -or (Read-Value $path 'Class') -cne $Class -or -not ([string](Read-Value $path 'Assembly')).StartsWith('VBAi,', [StringComparison]::Ordinal)) { throw "Foreign COM server/version ownership: $path" }
        }
        if (-not (Read-Value $path 'CodeBase')) { throw "Missing CodeBase: $path" }
    }
    return $paths
}

try {
    if ($Restore) {
        $recordPath = (Resolve-Path -LiteralPath $BackupPath).Path
        $beforeRecord = Import-Clixml -LiteralPath $recordPath
        if (-not (Test-Path -LiteralPath ($recordPath + '.after.clixml'))) { throw 'Applied-state companion is missing; inspect the immutable prior-state backup before any restoration.' }
        $record = Import-Clixml -LiteralPath ($recordPath + '.after.clixml')
        if ($beforeRecord.Format -cne 'VBAi.TestExplorerCandidate.1' -or $beforeRecord.BeforeFingerprint -cne $record.BeforeFingerprint -or (Fingerprint $beforeRecord.Before) -cne $record.BeforeFingerprint) { throw 'Immutable prior-state backup does not match the applied-state record.' }
        if ($record.Format -cne 'VBAi.TestExplorerCandidate.1' -or $record.Hive -cne 'HKCU' -or $record.View -cne 'Registry64' -or $null -eq $record.After) { throw 'Backup is not a completed candidate registration snapshot.' }
        if (-not $record.AfterIsOwned -or -not (Owned-Difference $record.Before $record.After $record.AllowedKeys $record.AllowedValues)) { throw 'Concurrent or unexpected changes occurred during Apply; automatic snapshot Restore is refused.' }
        if ((@($record.Before | ForEach-Object Path) -join '|') -cne ($roots -join '|') -or (@($record.After | ForEach-Object Path) -join '|') -cne ($roots -join '|')) { throw 'Backup root scope is invalid.' }
        for ($index = 0; $index -lt $roots.Count; $index++) { Require-TreeScope $record.Before[$index] $roots[$index]; Require-TreeScope $record.After[$index] $roots[$index] }
        if ((Fingerprint $record.Before) -cne $record.BeforeFingerprint -or (Fingerprint $record.After) -cne $record.AfterFingerprint) { throw 'Backup snapshot integrity check failed.' }
        $current = Read-All
        if ((Fingerprint $current) -cne $record.AfterFingerprint) { throw 'Registration changed after candidate application. Restore refused to preserve concurrent changes.' }
        # Refuse restoration while this candidate is still loaded; never unload it automatically.
        $loaded = @()
        foreach ($process in @(Get-Process -Name EXCEL,WINWORD,POWERPNT,MSACCESS,MSPUB,OUTLOOK,SLDWORKS -ErrorAction SilentlyContinue)) {
            try { foreach ($module in $process.Modules) { if ($module.FileName -ieq $record.Candidate) { $loaded += $process.Id } } }
            catch { throw "Cannot verify loaded candidate in PID $($process.Id); restore refused." }
            finally { $process.Dispose() }
        }
        if ($loaded.Count) { throw "Candidate still loaded in PID(s) $($loaded -join ', '); inspect and close only owned qualification hosts before explicit Restore." }
        # Recheck immediately before the first write.
        if ((Fingerprint (Read-All)) -cne $record.AfterFingerprint) { throw 'Concurrent registration change before restore; no restore was attempted.' }
        for ($index = 0; $index -lt $roots.Count; $index++) {
            if ((Fingerprint (Read-Tree $roots[$index])) -cne (Fingerprint $record.After[$index])) { throw 'Concurrent registration change during restore; remaining keys were not restored. Preserve backup and inspect partial restoration.' }
            $registry.DeleteSubKeyTree($roots[$index], $false)
            Write-Tree $record.Before[$index]
        }
        if ((Fingerprint (Read-All)) -cne $record.BeforeFingerprint) { throw 'Restoration did not match the exact prior snapshot. Preserve the backup and inspect registry; no automatic retry.' }
        [pscustomobject]@{ Restored = $true; Verified = $true; Backup = $recordPath; Hive = 'HKCU'; View = 'Registry64' }
        return
    }

    $candidate = (Resolve-Path -LiteralPath $CandidateAssemblyPath).Path
    $assembly = [Reflection.Assembly]::ReflectionOnlyLoadFrom($candidate)
    if ($assembly.GetName().Name -cne 'VBAi' -or $assembly.ManifestModule.ModuleVersionId -ne $ExpectedMvid) { throw 'Candidate assembly identity/MVID mismatch.' }
    $candidateFolder = Split-Path -Parent $candidate
    $dependencies = @('Markdig.dll','System.Resources.Extensions.dll','System.Memory.dll','System.Buffers.dll','System.Numerics.Vectors.dll','System.Runtime.CompilerServices.Unsafe.dll','Microsoft.Web.WebView2.Core.dll','Microsoft.Web.WebView2.WinForms.dll','VBAi.Updater.exe')
    foreach ($dependency in $dependencies) { if (-not (Test-Path -LiteralPath (Join-Path $candidateFolder $dependency) -PathType Leaf)) { throw "Candidate dependency missing: $dependency" } }
    if (-not (Test-Path -LiteralPath (Join-Path $candidateFolder 'EditorAssets\index.html') -PathType Leaf)) { throw 'Candidate EditorAssets/index.html is missing.' }
    if (-not (Test-Path -LiteralPath (Join-Path $candidateFolder 'runtimes\win-x64\native\WebView2Loader.dll') -PathType Leaf) -and -not (Test-Path -LiteralPath (Join-Path $candidateFolder 'WebView2Loader.dll') -PathType Leaf)) { throw 'Candidate x64 WebView2Loader.dll is missing.' }
    $paths = @((Require-Server $addIn 'VBAi.AddIn' 'VBAi.AddIn' '{8E854243-087F-4D6C-9E0E-8622B0E50883}')) + @((Require-Server $chat 'VBAi.ChatToolWindow' 'VBAi.ChatToolWindow' '{0F4D723B-97D8-42E5-9B31-70646B97C8D2}'))
    foreach ($path in $paths) {
        $installed = ([Uri][string](Read-Value $path 'CodeBase')).LocalPath
        if ([IO.Path]::GetFullPath($installed) -ieq $candidate) { throw 'Candidate path must differ from every currently registered AddIn/Chat DLL path.' }
        if ([IO.Path]::GetFullPath((Split-Path -Parent $installed)) -ieq [IO.Path]::GetFullPath($candidateFolder)) { throw 'Candidate must use a separate output directory, not the installed DLL directory.' }
        if ((Read-Value $path 'Assembly') -cne $assembly.FullName) { throw 'Candidate assembly identity differs from the existing server; CodeBase-only activation would be invalid.' }
    }
    $runtimeScript = Join-Path (Split-Path -Parent $PSScriptRoot) 'Register-VbaTestRuntime.ps1'
    $runtimePlan = & $runtimeScript -AssemblyPath $candidate -Preview
    $before = Read-All
    $candidateCodeBase = 'file:///' + $candidate.Replace([char]92, [char]47)
    $changes = @($paths | ForEach-Object { [pscustomobject]@{ Path = $_; Name = 'CodeBase'; Before = Read-Value $_ 'CodeBase'; After = $candidateCodeBase } })
    $allowedKeys = @{}; $allowedValues = @{}
    foreach ($change in $changes) { $allowedValues['V|' + $change.Path + [char]0 + 'CodeBase'] = ([pscustomobject]@{ Kind = 'String'; Data = $candidateCodeBase } | ConvertTo-Json -Compress) }
    foreach ($path in $runtimePlan.Proposed.Keys) {
        $parent = [string]$path
        while ($parent -ne 'Software\Classes' -and $parent -ne 'Software\Classes\CLSID') { $allowedKeys['K|' + $parent] = $true; $parent = $parent.Substring(0, $parent.LastIndexOf('\')) }
        foreach ($name in $runtimePlan.Proposed[$path].Keys) { $allowedValues['V|' + $path + [char]0 + $name] = ([pscustomobject]@{ Kind = 'String'; Data = $runtimePlan.Proposed[$path][$name] } | ConvertTo-Json -Compress) }
    }
    $hosts = @()
    $candidateLoaded = @()
    foreach ($process in @(Get-Process -Name EXCEL,WINWORD,POWERPNT,MSACCESS,MSPUB,OUTLOOK,SLDWORKS -ErrorAction SilentlyContinue)) {
        try {
            $modules = @($process.Modules | Where-Object { $_.ModuleName -ieq 'VBAi.dll' } | ForEach-Object FileName)
            $hosts += [pscustomobject]@{ Process = $process.ProcessName; Pid = $process.Id; LoadedVbai = $modules; Effect = 'Existing loaded assembly retained; future activation uses candidate.' }
            if ($modules -icontains $candidate) { $candidateLoaded += $process.Id }
        } catch { if ($Apply) { throw "Cannot inspect loaded VBAi in PID $($process.Id); Apply refused." }; $hosts += [pscustomobject]@{ Process = $process.ProcessName; Pid = $process.Id; InspectionError = $_.Exception.Message } }
        finally { $process.Dispose() }
    }
    $preview = [pscustomobject]@{ Mode = 'Preview'; Hive = 'HKCU'; View = 'Registry64'; Candidate = $candidate; Mvid = $ExpectedMvid.ToString('D'); CandidateSha256 = (Get-FileHash -LiteralPath $candidate -Algorithm SHA256).Hash; CodeBaseChanges = $changes; RuntimeRegistration = $runtimePlan; ExistingHosts = $hosts; FilesCopied = $false; TrustChanged = $false; ApplyRequires = 'Explicit -Apply after reviewing this concrete candidate and future-host scope.' }
    if (-not $Apply) { $preview; return }
    if ($candidateLoaded.Count) { throw "Candidate already loaded in PID(s) $($candidateLoaded -join ', '); Apply refused." }
    $backup = if ($ReportPath) { [IO.Path]::GetFullPath($ReportPath) } else { Join-Path $candidateFolder ('candidate-registration-' + [Guid]::NewGuid().ToString('N') + '.clixml') }
    if (Test-Path -LiteralPath $backup) { throw 'Backup path already exists; overwrite refused.' }
    if ((Test-Path -LiteralPath ($backup + '.after.clixml')) -or (Test-Path -LiteralPath ($backup + '.runtime-before.clixml'))) { throw 'Backup companion path already exists; overwrite refused.' }
    [IO.Directory]::CreateDirectory((Split-Path -Parent $backup)) | Out-Null
    $record = [pscustomobject]@{ Format = 'VBAi.TestExplorerCandidate.1'; Hive = 'HKCU'; View = 'Registry64'; Candidate = $candidate; Mvid = $ExpectedMvid.ToString('D'); Before = $before; BeforeFingerprint = Fingerprint $before; After = $null; AfterFingerprint = $null; AfterIsOwned = $false; AllowedKeys = $allowedKeys; AllowedValues = $allowedValues; Phase = 'Prepared'; Preview = $preview }
    $record | Export-Clixml -LiteralPath $backup -Depth 100
    if ((Fingerprint (Read-All)) -cne $record.BeforeFingerprint) { throw 'Concurrent registration change before Apply; candidate was not applied.' }
    try {
        foreach ($path in $paths) { $key = $registry.OpenSubKey($path, $true); try { $key.SetValue('CodeBase', $candidateCodeBase, [Microsoft.Win32.RegistryValueKind]::String) } finally { $key.Dispose() } }
        & $runtimeScript -AssemblyPath $candidate -ReportPath ($backup + '.runtime-before.clixml') | Out-Null
        foreach ($path in $paths) { if ((Read-Value $path 'CodeBase') -cne $candidateCodeBase) { throw 'Candidate CodeBase readback mismatch.' } }
        $record.Phase = 'Applied'
    } catch { $record.Phase = 'PartialOrFailed'; throw "Candidate registration may be partial. Backup: $backup. Explicit Restore required after inspection; no automatic rollback. $($_.Exception.Message)" }
    finally {
        $record.After = Read-All; $record.AfterFingerprint = Fingerprint $record.After
        $record.AfterIsOwned = Owned-Difference $record.Before $record.After $allowedKeys $allowedValues
        # Keep the prior-state backup immutable, including after a partial mutation.
        $record | Export-Clixml -LiteralPath ($backup + '.after.clixml') -Depth 100
        if (-not $record.AfterIsOwned) { throw "Unexpected concurrent registry changes during Apply. Preserved before/after snapshots: $backup. Explicit Restore will refuse to overwrite them." }
    }
    [pscustomobject]@{ Applied = $true; Backup = $backup; Candidate = $candidate; Mvid = $ExpectedMvid.ToString('D'); ExistingHosts = $hosts; Next = 'Qualify only new owned host processes. Restore explicitly after their verified normal exit.' }
} finally { $registry.Dispose() }
