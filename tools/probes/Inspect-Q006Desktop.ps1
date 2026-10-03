#requires -Version 5.1
param(
    [Parameter(Mandatory=$true)][string]$Desktop,
    [Parameter(Mandatory=$true)][int]$OwnedPid,
    [Parameter(Mandatory=$true)][string]$ExpectedStartUtc,
    [Parameter(Mandatory=$true)][string]$OutputPath,
    [long]$DialogHandle,
    [string]$HelperAssembly,
    [switch]$PrivateWorker,
    [switch]$CloseTerminalStartupDialog,
    [switch]$CloseOwnedSeedHost,
    [string]$ExpectedSeed,
    [string]$ProductAssembly
)
$ErrorActionPreference='Stop'
if ($Desktop -notmatch '^VBAiTests_[0-9a-f]{32}$' -or (Test-Path -LiteralPath $OutputPath)) { throw 'Exact generated desktop and new evidence path required.' }
$owned = Get-Process -Id $OwnedPid -ErrorAction Stop
if ($owned.ProcessName -ne 'EXCEL' -or $owned.StartTime.ToUniversalTime().ToString('o') -cne $ExpectedStartUtc) { throw 'Owned Excel identity differs.' }
if ($DialogHandle -ne 0 -or $CloseOwnedSeedHost) {
    $type=[Reflection.Assembly]::LoadFrom($HelperAssembly).GetType('VBAi.Tests.Integration.IsolatedTestDesktop',$true)
    $flags=[Reflection.BindingFlags]'Static,NonPublic'
    if($PrivateWorker){ $type.GetMethod('RequireCurrent',$flags).Invoke($null,@($Desktop)) | Out-Null }
    else {
        $powershell=Join-Path $env:WINDIR 'System32/WindowsPowerShell/v1.0/powershell.exe'
        $arguments=[string[]]@('-NoProfile','-NonInteractive','-Sta','-File',$PSCommandPath,'-Desktop',$Desktop,'-OwnedPid',[string]$OwnedPid,
            '-ExpectedStartUtc',$ExpectedStartUtc,'-OutputPath',$OutputPath,'-DialogHandle',[string]$DialogHandle,'-HelperAssembly',$HelperAssembly,'-PrivateWorker')
        $parameters=New-Object object[] 4
        $parameters[0]=[string]$powershell; $parameters[1]=[string[]]$arguments
        $parameters[2]=[string]$PSScriptRoot; $parameters[3]=[string]$Desktop
        if($CloseTerminalStartupDialog){ $arguments += '-CloseTerminalStartupDialog'; $parameters[1]=[string[]]$arguments }
        if($CloseOwnedSeedHost){ $arguments += @('-CloseOwnedSeedHost','-ExpectedSeed',$ExpectedSeed,'-ProductAssembly',$ProductAssembly); $parameters[1]=[string[]]$arguments }
        $child=$type.GetMethod('Launch',$flags).Invoke($null,$parameters)
        $instance=[Reflection.BindingFlags]'Instance,NonPublic'
        if(-not $child.GetType().GetMethod('Wait',$instance).Invoke($child,@(30000))){throw 'Read-only observer is pending; do not replay or terminate it.'}
        $code=$child.GetType().GetMethod('ExitCode',$instance).Invoke($child,@())
        $child.Dispose()
        if($code -ne 0){throw "Private read-only observer failed: $code"}
        Get-Content -LiteralPath $OutputPath -Raw -Encoding UTF8
        exit 0
    }
}
if($CloseOwnedSeedHost){
    $product=[Reflection.Assembly]::LoadFrom($ProductAssembly)
    $resolver=$product.GetType('VBAi.ExcelOwnedApplication',$true).GetMethod('Resolve',[Reflection.BindingFlags]'Static,NonPublic')
    $fallback=[Func[object]]{throw 'Exact owned native document is unavailable; no activation fallback.'}
    $application=$resolver.Invoke($null,@([int]$OwnedPid,$fallback))
    $books=$application.Workbooks
    if($books.Count -ne 1){throw 'Cleanup requires only the exact macro-free seed.'}
    $book=$books.Item(1)
    if($book.FullName -cne $ExpectedSeed -or -not $book.Saved -or $book.HasVBProject){throw 'Owned seed identity or inert state differs; no Quit.'}
    @{State='SINGLE_NORMAL_QUIT_INTENT';OwnedPid=$OwnedPid;Seed=$ExpectedSeed;Saved=$book.Saved;HasVBProject=$book.HasVBProject;Desktop=$Desktop;Utc=[DateTime]::UtcNow.ToString('o')} |
        ConvertTo-Json | Set-Content -LiteralPath ($OutputPath+'.intent.json') -Encoding UTF8
    $application.Quit()
    [Runtime.InteropServices.Marshal]::FinalReleaseComObject($book) | Out-Null
    [Runtime.InteropServices.Marshal]::FinalReleaseComObject($books) | Out-Null
    [Runtime.InteropServices.Marshal]::FinalReleaseComObject($application) | Out-Null
    $exited=$owned.WaitForExit(10000)
    $exitCode=if($exited){$owned.ExitCode}else{$null}
    @{State='NORMAL_QUIT_OBSERVED';OwnedPid=$OwnedPid;Exited=$exited;ExitCode=$exitCode;ForcedTermination=$false;Seed=$ExpectedSeed;Utc=[DateTime]::UtcNow.ToString('o')} |
        ConvertTo-Json | Set-Content -LiteralPath $OutputPath -Encoding UTF8
    if(-not $exited -or $exitCode -ne 0){throw 'Owned cleanup did not observe a normal exit; preserve evidence.'}
    exit 0
}
Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Text;
public static class Q006DesktopInventory {
    public sealed class Row { public long Handle, Parent; public uint Pid, Thread; public string Class, Caption; public bool Visible, Enabled; }
    private delegate bool Visitor(IntPtr handle, IntPtr state);
    [DllImport("user32.dll", CharSet=CharSet.Unicode, SetLastError=true)] private static extern IntPtr OpenDesktopW(string name, uint flags, bool inherit, uint access);
    [DllImport("user32.dll")] private static extern bool CloseDesktop(IntPtr handle);
    [DllImport("user32.dll", SetLastError=true)] private static extern bool EnumDesktopWindows(IntPtr handle, Visitor visitor, IntPtr state);
    [DllImport("user32.dll")] private static extern bool EnumChildWindows(IntPtr handle, Visitor visitor, IntPtr state);
    [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr handle, out uint pid);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] private static extern int GetClassNameW(IntPtr handle, StringBuilder text, int count);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] private static extern int GetWindowTextW(IntPtr handle, StringBuilder text, int count);
    [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr handle);
    [DllImport("user32.dll")] private static extern bool IsWindowEnabled(IntPtr handle);
    [DllImport("user32.dll", SetLastError=true)] private static extern bool PostMessageW(IntPtr handle, uint message, IntPtr wParam, IntPtr lParam);
    public static void CloseWindow(long handle, uint expectedPid) {
        uint pid; GetWindowThreadProcessId(new IntPtr(handle),out pid);
        if(pid!=expectedPid || !PostMessageW(new IntPtr(handle),0x0010,IntPtr.Zero,IntPtr.Zero)) throw new InvalidOperationException("Exact owned normal close could not be delivered.");
    }
    public static Row[] Read(string desktop, uint expectedPid) {
        IntPtr scope=OpenDesktopW(desktop, 0, false, 0x41);
        if(scope==IntPtr.Zero) throw new Win32Exception(Marshal.GetLastWin32Error());
        var rows=new List<Row>();
        Action<IntPtr,IntPtr> add=(window,parent)=> {
            uint pid; uint thread=GetWindowThreadProcessId(window,out pid); if(pid!=expectedPid) return;
            if(rows.Count>=4096) throw new InvalidOperationException("Inventory bound exceeded.");
            var name=new StringBuilder(256); var title=new StringBuilder(512);
            GetClassNameW(window,name,name.Capacity); GetWindowTextW(window,title,title.Capacity);
            rows.Add(new Row{Handle=window.ToInt64(),Parent=parent.ToInt64(),Pid=pid,Thread=thread,
                Class=name.ToString(),Caption=title.ToString(),Visible=IsWindowVisible(window),Enabled=IsWindowEnabled(window)});
        };
        try {
            Visitor roots=(window,state)=> {
                uint pid; GetWindowThreadProcessId(window,out pid); if(pid!=expectedPid) return true;
                add(window,IntPtr.Zero); Visitor children=(child,s)=>{add(child,window);return true;};
                EnumChildWindows(window,children,IntPtr.Zero); return true;
            };
            if(!EnumDesktopWindows(scope,roots,IntPtr.Zero)) throw new Win32Exception(Marshal.GetLastWin32Error());
            return rows.ToArray();
        } finally {CloseDesktop(scope);}
    }
}
'@
$rows=[Q006DesktopInventory]::Read($Desktop,[uint32]$OwnedPid)
$uia=@()
if($DialogHandle -ne 0){
    if(-not @($rows | Where-Object {$_.Handle -eq $DialogHandle -and $_.Pid -eq $OwnedPid -and $_.Parent -eq 0}).Count){throw 'The exact owned top-level dialog was not observed.'}
    Add-Type -AssemblyName UIAutomationClient,UIAutomationTypes
    $element=[Windows.Automation.AutomationElement]::FromHandle([IntPtr]$DialogHandle)
    $children=$element.FindAll([Windows.Automation.TreeScope]::Descendants,[Windows.Automation.Condition]::TrueCondition)
    if($children.Count -gt 512){throw 'UIA observation exceeds its bound.'}
    foreach($child in $children){
        $current=$child.Current
        if($current.ProcessId -ne $OwnedPid){throw 'UIA crossed the owned process boundary.'}
        $uia+=@{Name=$current.Name;Id=$current.AutomationId;Type=$current.ControlType.ProgrammaticName;Enabled=$current.IsEnabled;Handle=$current.NativeWindowHandle}
    }
    if($CloseTerminalStartupDialog){
        if(-not @($uia | Where-Object {$_.Type -eq 'ControlType.Text' -and ($_.Name -like '*/x.xlsx*' -or $_.Name -like '*/automation.xlsx*')}).Count){throw 'Only the observed terminal synthetic-switch file error may be closed.'}
        [Q006DesktopInventory]::CloseWindow($DialogHandle,[uint32]$OwnedPid)
    }
}
@{Desktop=$Desktop;OwnedPid=$OwnedPid;StartUtc=$ExpectedStartUtc;ReadOnly=(-not $CloseTerminalStartupDialog);NormalDialogClose=$CloseTerminalStartupDialog.IsPresent;Windows=$rows;UiA=$uia;Utc=[DateTime]::UtcNow.ToString('o')} |
    ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $OutputPath -Encoding UTF8
$rows | Where-Object {$_.Parent -eq 0 -or $_.Class -eq 'EXCEL7'} | Select-Object Handle,Thread,Class,Caption,Visible,Enabled
if($uia.Count){$uia | ConvertTo-Json -Depth 4}
