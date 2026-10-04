#requires -Version 5.1
param([Parameter(Mandatory=$true)][string]$EvidenceRoot)
$ErrorActionPreference='Stop'
if($PSVersionTable.PSEdition -ne 'Desktop'){throw 'Desktop PowerShell required'}
$output=Join-Path $EvidenceRoot 'recovery-restoration.json'
if(Test-Path -LiteralPath $output){throw 'Recovery restoration is one-shot'}
$shutdown=Get-Content (Join-Path $EvidenceRoot 'bootstrap-shutdown.json') -Raw -Encoding UTF8|ConvertFrom-Json
if($shutdown.State -ne 'BOOTSTRAP_FAILURE_HOST_NORMAL_EXIT' -or $shutdown.ExitCode -ne 0){throw 'Owned host normal-exit receipt required'}
if(@(Get-Process EXCEL,WINWORD,POWERPNT,MSACCESS,MSPUB,OUTLOOK -ErrorAction SilentlyContinue).Count){throw 'Office hosts remain; restoration refused'}
$plan=Get-Content (Join-Path $EvidenceRoot 'q028-plan.json') -Raw -Encoding UTF8|ConvertFrom-Json
$intent=Get-Content (Join-Path $EvidenceRoot 'settings-intent.json') -Raw -Encoding UTF8|ConvertFrom-Json
if((Get-FileHash -LiteralPath $plan.Product).Hash -cne $plan.ProductSha256){throw 'Product bytes changed'}
$assembly=[Reflection.Assembly]::LoadFrom($plan.Product)
$type=$assembly.GetType('VBAi.LlmSettings',$true)
$load=$type.GetMethod('Load');$save=$type.GetMethod('Save');$settings=$load.Invoke($null,@())
$matches=@($intent.OnlyChangedFields|ForEach-Object {
    $value=$type.GetProperty($_).GetValue($settings,$null)
    @{Field=$_;Applied=($value -ceq $intent.Applied.$_);Baseline=($value -ceq $intent.Baseline.$_)}
})
@{SettingsPath=$type.GetProperty('FilePath',[Reflection.BindingFlags]'Static,NonPublic').GetValue($null,$null);Matches=$matches}|ConvertTo-Json -Depth 5|Set-Content -LiteralPath (Join-Path $EvidenceRoot 'restoration-preflight.json') -Encoding UTF8
$alreadyBaseline=@($matches|Where-Object {-not $_.Baseline}).Count -eq 0
foreach($field in $intent.OnlyChangedFields){
    if($alreadyBaseline){continue}
    if($type.GetProperty($field).GetValue($settings,$null) -cne $intent.Applied.$field){throw "Concurrent settings change: $field"}
}
$receipt=@{Qualified=$false;State='RESTORE_INTENT';OnlyChangedFields=$intent.OnlyChangedFields;SecretsCaptured=$false;SettingsSaveEntries=0;AlreadyBaseline=$alreadyBaseline}
$receipt|ConvertTo-Json -Depth 8|Set-Content -LiteralPath ($output+'.progress.json') -Encoding UTF8
if(-not $alreadyBaseline){
    foreach($field in $intent.OnlyChangedFields){$type.GetProperty($field).SetValue($settings,$intent.Baseline.$field,$null)}
    $receipt.SettingsSaveEntries=1
    $save.Invoke($settings,@())|Out-Null
}
$readback=$load.Invoke($null,@())
foreach($field in $intent.OnlyChangedFields){if($type.GetProperty($field).GetValue($readback,$null) -cne $intent.Baseline.$field){throw "Restoration mismatch: $field"}}
$receipt.SettingsRestored=$true
try {
    $receipt.RegistrationRestore=& (Join-Path $plan.Repository 'tools/testing-explorer/Set-TestExplorerCandidate.ps1') -Restore -BackupPath (Join-Path $EvidenceRoot 'candidate-registration.clixml')
} catch {
    $receipt.State='REGISTRATION_RESTORE_REFUSED_CONCURRENT_CHANGE';$receipt.RegistrationRestoreError=$_.Exception.Message
    $receipt.Utc=[DateTime]::UtcNow.ToString('o')
    $receipt|ConvertTo-Json -Depth 8|Set-Content -LiteralPath $output -Encoding UTF8
    throw
}
$receipt.State='RESTORED_AFTER_FAILED_BOOTSTRAP';$receipt.Utc=[DateTime]::UtcNow.ToString('o')
$receipt|ConvertTo-Json -Depth 8|Set-Content -LiteralPath $output -Encoding UTF8
Get-Content -LiteralPath $output -Raw -Encoding UTF8
