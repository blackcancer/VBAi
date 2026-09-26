using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

namespace CodexVBE
{
    [ComImport, Guid("7BF80981-BF32-101A-8BBB-00AA00300CAB"),
     InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    internal interface IOlePictureDisp
    {
        [DispId(0)] int Handle { get; }
        [DispId(3)] short Type { get; }
        [DispId(4)] int Width { get; }
        [DispId(5)] int Height { get; }
    }

    internal static class OlePictureLoader
    {
        [DllImport("oleaut32.dll", EntryPoint = "OleLoadPictureFile", PreserveSig = false)]
        private static extern void OleLoadPictureFile(
            [MarshalAs(UnmanagedType.Struct)] object fileName,
            [MarshalAs(UnmanagedType.IDispatch)] out object picture);

        public static object Load(string path)
        {
            if (string.IsNullOrWhiteSpace(path) ||
                !Regex.IsMatch(path, @"^(?:[A-Za-z]:[\\/]|\\\\[^\\/]+[\\/][^\\/]+[\\/])"))
                throw new ArgumentException("Picture path must be a fully qualified local or UNC path.");
            string fullPath = System.IO.Path.GetFullPath(path);
            if (!File.Exists(fullPath)) throw new FileNotFoundException("Picture file not found.", fullPath);
            OleLoadPictureFile(fullPath, out object picture);
            if (picture == null || !Marshal.IsComObject(picture))
                throw new InvalidOperationException("Windows did not load an OLE picture from the supplied file.");
            return picture;
        }

        public static string Fingerprint(object picture)
        {
            if (picture == null) return null;
            var typed = picture as IOlePictureDisp;
            if (typed == null)
                throw new InvalidOperationException("The UserForm Picture is not an OLE picture.");
            return typed.Type + ":" + typed.Width + ":" + typed.Height + ":" + typed.Handle;
        }
    }
}
