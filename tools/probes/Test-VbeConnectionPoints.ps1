param(
    [Parameter(Mandatory = $true)][string]$AssemblyPath,
    [Parameter(Mandatory = $true)][string]$OutputDirectory,
    [switch]$AllowTemporaryVbaAccess
)
$ErrorActionPreference = 'Stop'
if (@(Get-Process EXCEL -ErrorAction SilentlyContinue).Count) { throw 'Close existing Excel instances before this isolated test.' }
$assembly = [Reflection.Assembly]::LoadFrom((Resolve-Path -LiteralPath $AssemblyPath))
$directory = [IO.Path]::GetFullPath($OutputDirectory)
[IO.Directory]::CreateDirectory($directory) | Out-Null
$document = Join-Path $directory ('Vbai-EditorProbe-' + [Guid]::NewGuid().ToString('N') + '.xlsm')
$securityPath = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Security'
$prior = Get-ItemProperty -LiteralPath $securityPath -ErrorAction Stop
$hadAccess = $null -ne $prior.PSObject.Properties['AccessVBOM']
$priorAccess = $prior.AccessVBOM
if ($priorAccess -ne 1 -and -not $AllowTemporaryVbaAccess) { throw 'Explicit -AllowTemporaryVbaAccess is required to temporarily enable trusted VBA access.' }
$probeProcess = $null
$otherBook = $null
$excel = $null; $book = $null; $form = $null; $session = $null
function Invoke-Session([hashtable]$Fields) {
    $request = New-Object CodexVBE.Request
    foreach ($key in $Fields.Keys) { $request.$key = $Fields[$key] }
    $response = $script:session.Execute($request)
    if (-not $response.Ok) { throw $response.Error }
    return $response.Data
}
try {
    if ($priorAccess -ne 1) { New-ItemProperty -LiteralPath $securityPath -Name AccessVBOM -PropertyType DWord -Value 1 -Force | Out-Null }
    $excel = New-Object -ComObject Excel.Application
    $probeProcess = Get-Process EXCEL -ErrorAction Stop
    if (@($probeProcess).Count -ne 1) { throw "Excel isolation was lost." }
    $excel.Visible = $true
    $book = $excel.Workbooks.Add()
    $vbe = $excel.GetType().InvokeMember('VBE', [Reflection.BindingFlags]::GetProperty, $null, $excel, $null)
    $vbe.MainWindow.Visible = $true
    $session = [Activator]::CreateInstance($assembly.GetType('CodexVBE.VbeSession', $true), [object[]]@($vbe))
    $project = $book.VBProject
    Add-Type -TypeDefinition @"
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices.ComTypes;
public static class VbeConnectionProbe {
    public static string[] Read(object target) {
        var result = new List<string>();
        var container = target as IConnectionPointContainer;
        if (container == null) return new[] { "IConnectionPointContainer unavailable" };
        foreach (var text in new[] { "0002E103-0000-0000-C000-000000000046", "0002E116-0000-0000-C000-000000000046", "CDDE3804-2064-11CF-867F-00AA005FF34A" }) {
            Guid iid = new Guid(text);
            try { IConnectionPoint point; container.FindConnectionPoint(ref iid, out point); if (point == null) result.Add(text + ": null"); else { Guid actual; point.GetConnectionInterface(out actual); result.Add(text + ": actual " + actual.ToString()); } }
            catch (Exception error) { result.Add(text + ": " + error.GetType().Name + " " + error.HResult.ToString("X8")); }
        }
        return result.ToArray();
    }
}
"@
    $result = @()
    foreach ($entry in @(@{ Name='VBE'; Source=$vbe }, @{ Name='Events'; Source=$vbe.Events }, @{ Name='VBProjects'; Source=$vbe.VBProjects }, @{ Name='VBProject'; Source=$project }, @{ Name='VBComponents'; Source=$project.VBComponents }, @{ Name='References'; Source=$project.References })) {
        try { $result += [pscustomobject]@{ Name=$entry.Name; Interfaces=[VbeConnectionProbe]::Read($entry.Source); Error=$null } }
        catch { $result += [pscustomobject]@{ Name=$entry.Name; Interfaces=@(); Error=$_.Exception.Message } }
    }
    $result | ConvertTo-Json -Depth 5 | Set-Content (Join-Path $directory 'vbe-event-connection-points.json') -Encoding UTF8
    $result | ConvertTo-Json -Depth 5
} finally {
    if ($hadAccess) { New-ItemProperty -LiteralPath $securityPath -Name AccessVBOM -PropertyType DWord -Value $priorAccess -Force | Out-Null }
    else { Remove-ItemProperty -LiteralPath $securityPath -Name AccessVBOM -ErrorAction SilentlyContinue }
    try { if ($null -ne $otherBook) { $otherBook.Close($false) }; if ($null -ne $book) { $book.Close($false) } }
    finally {
        try { if ($null -ne $excel) { $excel.Quit() } }
        finally { if ($null -ne $probeProcess -and -not $probeProcess.WaitForExit(2000)) { Stop-Process -Id $probeProcess.Id -Force } }
    }
}
