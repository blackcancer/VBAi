using System;
using System.Runtime.InteropServices;

namespace CodexVBE
{
    // The interface IDs and member order come from the installed VBE Extensibility type library.
    // CreateToolWindow requires the typed AddIn interface; late-bound object/IDispatch
    // arguments are rejected by VBE with DISP_E_TYPEMISMATCH.
    [ComImport]
    [Guid("DA936B64-AC8B-11D1-B6E5-00A0C90F2744")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    internal interface IVbeAddIn { }

    [ComImport]
    [Guid("F57B7ED0-D8AB-11D1-85DF-00C04F98F42C")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    internal interface IVbeWindows
    {
        [DispId(1)] object VBE { [return: MarshalAs(UnmanagedType.Interface)] get; }
        [DispId(2)] object Parent { [return: MarshalAs(UnmanagedType.Interface)] get; }
        [DispId(0)] [return: MarshalAs(UnmanagedType.Interface)] object Item(
            [In, MarshalAs(UnmanagedType.Struct)] object index);
        [DispId(201)] int Count { get; }
        [DispId(-4)] System.Collections.IEnumerator GetEnumerator();

        [DispId(300)]
        [return: MarshalAs(UnmanagedType.Interface)]
        object CreateToolWindow(
            [In, MarshalAs(UnmanagedType.Interface)] IVbeAddIn addIn,
            [In, MarshalAs(UnmanagedType.BStr)] string progId,
            [In, MarshalAs(UnmanagedType.BStr)] string caption,
            [In, MarshalAs(UnmanagedType.BStr)] string position,
            [In, Out, MarshalAs(UnmanagedType.Struct)] ref object document);
    }
}
