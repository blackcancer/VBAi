param([string]$OutputDirectory = 'artifacts/designer-compatibility/metadata')
$ErrorActionPreference = 'Stop'
$repository = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
$proof = @()
function Read-Items($project) {
    $json = & dotnet msbuild $project -nologo -p:DesignTimeBuild=true '-getItem:Compile,EmbeddedResource'
    if ($LASTEXITCODE -ne 0) { throw "Cannot evaluate $project" }
    return ($json -join "`n" | ConvertFrom-Json).Items
}
function Require-Item($items, $identity, $subtype, $parent) {
    $entry = @($items | Where-Object { $_.Identity -eq $identity })
    if ($entry.Count -ne 1) { throw "Missing or duplicate item $identity" }
    if ($subtype -and $entry[0].SubType -ne $subtype) { throw "$identity must have SubType=$subtype" }
    if ($parent -and $entry[0].DependentUpon -ne $parent) { throw "$identity must depend on $parent" }
    return [pscustomobject]@{ File = $identity; SubType = $entry[0].SubType; Parent = $entry[0].DependentUpon; Result = 'PASS' }
}
$items = Read-Items (Join-Path $repository 'src/CodexVBE/CodexVBE.csproj')
foreach ($name in @('Editor\ModernEditorWindow','Updates\UpdateWindow','Updates\UpdateProgressWindow','Git\GitWindow','Llm\Settings\LlmSettingsWindow')) {
    $parent = ($name -split '\\')[-1] + '.cs'
    $proof += Require-Item $items.Compile ($name+'.cs') 'Form' $null
    $proof += Require-Item $items.Compile ($name+'.Designer.cs') 'Code' $parent
    $proof += Require-Item $items.EmbeddedResource ($name+'.resx') $null $parent
}
foreach ($entry in $items.Compile | Where-Object { $_.Identity -match '(ModernEditorWindow|GitWindow|LlmSettingsWindow)\.[^.]+\.cs$' }) {
    $parent = [regex]::Match($entry.Identity, '(ModernEditorWindow|GitWindow|LlmSettingsWindow)\.').Groups[1].Value + '.cs'
    $proof += Require-Item $items.Compile $entry.Identity 'Code' $parent
}
$updater = Read-Items (Join-Path $repository 'src/VBAi.Updater/VBAi.Updater.csproj')
$proof += Require-Item $updater.Compile '../CodexVBE/Updates/UpdateProgressWindow.cs' 'Code' $null
$proof += Require-Item $updater.Compile '../CodexVBE/Updates/UpdateProgressWindow.Designer.cs' 'Code' 'UpdateProgressWindow.cs'
$directory = [IO.Path]::GetFullPath((Join-Path $repository $OutputDirectory))
[IO.Directory]::CreateDirectory($directory) | Out-Null
$proof | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $directory 'project-items.json') -Encoding UTF8
Write-Output "PASS $($proof.Count) evaluated WinForms project items"
