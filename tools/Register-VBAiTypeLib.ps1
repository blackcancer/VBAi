param([switch] $Unregister)

$ErrorActionPreference = 'Stop'
if (-not [Environment]::Is64BitProcess) { throw 'Use 64-bit PowerShell.' }

$root = Split-Path -Parent $PSScriptRoot
$typeLibPath = Join-Path $root 'bin\Debug\net48\VBAi.tlb'
if (-not (Test-Path -LiteralPath $typeLibPath)) { throw "Missing type library: $typeLibPath" }
$typeLibPath = (Resolve-Path -LiteralPath $typeLibPath).Path

Add-Type @'
using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
public static class CodexTypeLibRegistration {
    [DllImport("oleaut32.dll", CharSet = CharSet.Unicode, PreserveSig = false)]
    public static extern void LoadTypeLibEx(string path, int registrationKind, out ITypeLib typeLib);
    [DllImport("oleaut32.dll", CharSet = CharSet.Unicode, PreserveSig = false)]
    public static extern void RegisterTypeLibForUser(ITypeLib typeLib, string path, string helpDirectory);
    [DllImport("oleaut32.dll", PreserveSig = false)]
    public static extern void UnRegisterTypeLibForUser(ref Guid typeLibId, ushort major, ushort minor, int locale, int systemKind);
    public static string Apply(string path, bool unregister) {
        ITypeLib lib;
        LoadTypeLibEx(path, 2, out lib);
        IntPtr attributesPointer = IntPtr.Zero;
        try {
            lib.GetLibAttr(out attributesPointer);
            System.Runtime.InteropServices.ComTypes.TYPELIBATTR attributes =
                (System.Runtime.InteropServices.ComTypes.TYPELIBATTR)Marshal.PtrToStructure(
                    attributesPointer, typeof(System.Runtime.InteropServices.ComTypes.TYPELIBATTR));
            if (unregister) {
                Guid id = attributes.guid;
                UnRegisterTypeLibForUser(ref id, (ushort)attributes.wMajorVerNum, (ushort)attributes.wMinorVerNum, attributes.lcid, (int)attributes.syskind);
            } else {
                RegisterTypeLibForUser(lib, path, System.IO.Path.GetDirectoryName(path));
            }
            return attributes.guid + " " + attributes.wMajorVerNum + "." + attributes.wMinorVerNum;
        } finally {
            if (attributesPointer != IntPtr.Zero) lib.ReleaseTLibAttr(attributesPointer);
            if (lib != null) Marshal.ReleaseComObject(lib);
        }
    }
}
'@

$identity = [CodexTypeLibRegistration]::Apply($typeLibPath, [bool]$Unregister)
if ($Unregister) { Write-Output "Unregistered user type library $identity" }
else { Write-Output "Registered user type library $identity" }
