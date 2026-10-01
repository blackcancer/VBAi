using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;

namespace VBAi.Tests.Integration
{
    internal sealed partial class ExcelVbeFixture
    {
        /// <summary>Loads synthetic observed StdFont descriptors into each imported font once, before owner transfer.</summary>
        internal void RestoreGitLayoutPersistedFonts(string form, string layout, IDictionary<string, object> expected)
        {
            WithGitLayoutDesigner(form, (component, designer) => {
                LoadGitFontOnce(designer, "Form.Font", expected);
                if (layout != "FrameMultiPage") return;
                object controls = null, frame = null;
                try
                {
                    controls = ((dynamic)designer).Controls;
                    frame = ((dynamic)controls).Item("QualificationExtra");
                    LoadGitFontOnce(frame, "Frame.Font", expected);
                }
                finally { Release(frame); Release(controls); }
            });
        }

        /// <summary>Transfers one bounded font descriptor without reading font members after Load.</summary>
        private static void LoadGitFontOnce(object owner, string prefix, IDictionary<string, object> expected)
        {
            byte[] name = Encoding.ASCII.GetBytes(Convert.ToString(expected[prefix + ".Name"]));
            decimal size = Convert.ToDecimal(expected[prefix + ".Size"]);
            if (name.Length == 0 || name.Length >= 32 || Encoding.ASCII.GetString(name) != Convert.ToString(expected[prefix + ".Name"]) ||
                size <= 0 || size > 200 || size * 10000m != decimal.Truncate(size * 10000m))
                throw new InvalidOperationException("Synthetic font descriptor is outside the exact diagnostic profile.");
            byte[] data;
            using (var memory = new MemoryStream())
            using (var writer = new BinaryWriter(memory))
            {
                writer.Write((byte)1);
                writer.Write(Convert.ToInt16(expected[prefix + ".Charset"]));
                writer.Write((byte)((Convert.ToBoolean(expected[prefix + ".Italic"]) ? 2 : 0) |
                    (Convert.ToBoolean(expected[prefix + ".Underline"]) ? 4 : 0) |
                    (Convert.ToBoolean(expected[prefix + ".Strikethrough"]) ? 8 : 0)));
                writer.Write(Convert.ToInt16(expected[prefix + ".Weight"]));
                writer.Write(checked((uint)(size * 10000m)));
                writer.Write((byte)name.Length);
                writer.Write(name);
                data = memory.ToArray();
            }
            object font = null;
            IStream stream = null;
            try
            {
                font = ((dynamic)owner).Font;
                var persisted = (GitDiagnosticPersistStream)font;
                Marshal.ThrowExceptionForHR(CreateGitFontStream(IntPtr.Zero, true, out stream));
                stream.Write(data, data.Length, IntPtr.Zero);
                stream.Seek(0, 0, IntPtr.Zero);
                persisted.Load(stream);
                ((dynamic)owner).Font = font;
            }
            finally { Release(stream); Release(font); }
        }

        /// <summary>Creates an owned in-memory OLE stream that frees its storage on final release.</summary>
        [DllImport("ole32.dll", EntryPoint = "CreateStreamOnHGlobal", ExactSpelling = true)]
        private static extern int CreateGitFontStream(IntPtr handle, [MarshalAs(UnmanagedType.Bool)] bool deleteOnRelease, out IStream stream);

        /// <summary>Standard persisted-font interface; this fixture loads only its fixed synthetic descriptors.</summary>
        [ComImport, Guid("00000109-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface GitDiagnosticPersistStream
        {
            void GetClassID(out Guid clsid);
            [PreserveSig] int IsDirty();
            void Load(IStream stream);
            void Save(IStream stream, [MarshalAs(UnmanagedType.Bool)] bool clearDirty);
            void GetSizeMax(out long size);
        }
    }
}
