namespace VBAi.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using System.Collections;
    using System.IO;
    using System.Runtime.InteropServices;
    using System.Runtime.InteropServices.ComTypes;
    using TYPELIBATTR = System.Runtime.InteropServices.ComTypes.TYPELIBATTR;
    using VBAi;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    public sealed partial class VbeReferenceTypesTests
    {
        private const string StdOlePath = @"C:\Windows\System32\stdole2.tlb";
        [DllImport("oleaut32.dll", CharSet = CharSet.Unicode, PreserveSig = false)]
        private static extern void LoadTypeLibEx(string path, int registrationKind, out ITypeLib library);
        private static FakeReference ReferenceFor(string path)
        {
            ITypeLib library;
            LoadTypeLibEx(path, 2, out library);
            IntPtr pointer = IntPtr.Zero;
            try
            {
                library.GetLibAttr(out pointer);
                var attr = (TYPELIBATTR)Marshal.PtrToStructure(pointer, typeof(TYPELIBATTR));
                return new FakeReference
                {
                    Name = "stdole2",
                    GUID = attr.guid.ToString("B"),
                    Major = attr.wMajorVerNum,
                    Minor = attr.wMinorVerNum,
                    FullPath = path
                };
            }
            finally
            {
                if (pointer != IntPtr.Zero)
                    library.ReleaseTLibAttr(pointer);
                Marshal.ReleaseComObject(library);
            }
        }

        private static object Prop(object value, string name)
        {
            return value.GetType().GetProperty(name).GetValue(value, null);
        }

        public sealed class FakeVbe
        {
            public List<FakeProject> VBProjects { get; } = new List<FakeProject>();
        }

        public sealed class FakeProject
        {
            public string Name { get; set; }
            public List<FakeReference> References { get; } = new List<FakeReference>();
        }

        public sealed class FakeReference
        {
            public string Name { get; set; }
            public string GUID { get; set; }
            public int Major { get; set; }
            public int Minor { get; set; }
            public bool IsBroken { get; set; }
            public string FullPath { get; set; }
        }
    }
}
