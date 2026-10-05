namespace VBAi.Tests.Unit
{
    using System;
    using System.Drawing;
    using System.Drawing.Imaging;
    using System.IO;
    using System.Runtime.InteropServices;
    using System.Runtime.InteropServices.ComTypes;
    using System.Threading;
    using VBAi;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    [TestClass]
    [TestCategory("Unit")]
    [DoNotParallelize]
    public sealed class OlePictureLoaderTests
    {
        [DataTestMethod]
        [DataRow("bmp", 1)]
        [DataRow("gif", 1)]
        [DataRow("jpg", 1)]
        [DataRow("ico", 3)]
        [DataRow("wmf", 2)]
        [DataRow("emf", 4)]
        public void NativeContentFingerprintSurvivesIndependentLoadAndPersistenceRoundtrip(string format, int expectedType)
        {
            OnSta(() =>
            {
                string directory = Path.Combine(Path.GetTempPath(), "vbai-ole-content-" + Guid.NewGuid().ToString("N"));
                Directory.CreateDirectory(directory);
                object first = null, second = null, changed = null, restored = null;
                try
                {
                    string original = Path.Combine(directory, "original." + format), different = Path.Combine(directory, "different." + format);
                    WriteOwnedPicture(original, format, Color.Red);
                    WriteOwnedPicture(different, format, Color.Blue);
                    first = OlePictureLoader.Load(original); second = OlePictureLoader.Load(original); changed = OlePictureLoader.Load(different);
                    var a = (IOlePictureDisp)first; var b = (IOlePictureDisp)second; var c = (IOlePictureDisp)changed;
                    Assert.AreEqual(expectedType, (int)a.Type, format);
                    Assert.AreNotEqual(a.Handle, b.Handle, "Independent native objects must actually have different GDI handles.");
                    Assert.AreEqual(a.Width, c.Width); Assert.AreEqual(a.Height, c.Height);
                    int originalHandle = a.Handle, originalDirty = ((IOlePicturePersistence)first).IsDirty();
                    byte[] originalMetafile = format == "wmf" ? ReadBorrowedMetafile(originalHandle) : null;
                    string fingerprint = OlePictureLoader.Fingerprint(first);
                    Assert.AreEqual(fingerprint, OlePictureLoader.Fingerprint(second), format + " independent load");
                    Assert.AreNotEqual(fingerprint, OlePictureLoader.Fingerprint(changed), format + " same-size changed content");
                    restored = PersistenceRoundtrip(first);
                    Assert.AreEqual(fingerprint, OlePictureLoader.Fingerprint(restored), format + " persistence roundtrip");
                    if (format == "wmf")
                    {
                        // Native OLE persistence can adapt the current WMF handle internally.
                        // Its previously returned borrowed resource must still be valid and unchanged.
                        CollectionAssert.AreEqual(originalMetafile, ReadBorrowedMetafile(originalHandle), "Previously borrowed WMF content changed or became invalid.");
                    }
                    else Assert.AreEqual(originalHandle, a.Handle, "Borrowed picture handle must remain intact.");
                    Assert.AreEqual(expectedType, (int)a.Type); Assert.AreEqual(b.Width, a.Width); Assert.AreEqual(b.Height, a.Height);
                    Assert.AreEqual(originalDirty, ((IOlePicturePersistence)first).IsDirty(), "Fingerprint must not clear dirty state.");
                    Assert.AreEqual(fingerprint, OlePictureLoader.Fingerprint(first), "Original object must remain usable after all reads.");
                }
                finally
                {
                    foreach (object owned in new[] { restored, changed, second, first })
                        if (owned != null && Marshal.IsComObject(owned)) Marshal.FinalReleaseComObject(owned);
                    Directory.Delete(directory, true);
                }
            });
        }

        [TestMethod]
        public void NativeEmptyPicturesHaveNoContentFingerprint()
        {
            OnSta(() =>
            {
                object uninitialized = null, none = null;
                IntPtr description = Marshal.AllocHGlobal(IntPtr.Size == 8 ? 24 : 20);
                try
                {
                    Guid iid = typeof(IOlePictureDisp).GUID;
                    OleCreatePictureIndirect(IntPtr.Zero, ref iid, true, out uninitialized);
                    Assert.AreEqual(-1, (int)((IOlePictureDisp)uninitialized).Type);
                    Assert.IsNull(OlePictureLoader.Fingerprint(uninitialized));
                    int size = IntPtr.Size == 8 ? 24 : 20;
                    Marshal.Copy(new byte[size], 0, description, size);
                    Marshal.WriteInt32(description, size); Marshal.WriteInt32(description, 4, 0);
                    OleCreatePictureIndirect(description, ref iid, true, out none);
                    Assert.AreEqual(0, (int)((IOlePictureDisp)none).Type);
                    Assert.IsNull(OlePictureLoader.Fingerprint(none));
                }
                finally
                {
                    if (none != null) Marshal.FinalReleaseComObject(none);
                    if (uninitialized != null) Marshal.FinalReleaseComObject(uninitialized);
                    Marshal.FreeHGlobal(description);
                }
            });
        }

        [TestMethod]
        public void ContentInspectionRefusesUnboundedOrChangingPersistenceWithoutReadingHandle()
        {
            var oversized = new PersistencePicture { Maximum = (ulong)OlePictureLoader.MaximumPersistenceBytes + 1 };
            Assert.ThrowsException<InvalidOperationException>(() => OlePictureLoader.Fingerprint(oversized));
            Assert.AreEqual(0, oversized.SaveCalls);
            var unavailableBound = new PersistencePicture { SizeError = unchecked((int)0x80004001) };
            Assert.ThrowsException<NotImplementedException>(() => OlePictureLoader.Fingerprint(unavailableBound));
            Assert.AreEqual(0, unavailableBound.SaveCalls, "Unsupported GetSizeMax must not silently allocate an unbounded stream.");
            var inaccurateBound = new PersistencePicture { Maximum = 1 };
            Assert.ThrowsException<InvalidOperationException>(() => OlePictureLoader.Fingerprint(inaccurateBound));
            Assert.AreEqual(1, inaccurateBound.SaveCalls);
            var changedDirty = new PersistencePicture { ChangeDirty = true };
            Assert.ThrowsException<InvalidOperationException>(() => OlePictureLoader.Fingerprint(changedDirty));
            Assert.AreEqual(1, changedDirty.SaveCalls);
            var unchanged = new PersistencePicture();
            Assert.AreEqual(OlePictureLoader.Fingerprint(unchanged), OlePictureLoader.Fingerprint(new PersistencePicture()));
            Assert.IsFalse(unchanged.ClearDirtyRequested);
        }

        private sealed class PersistencePicture : IOlePictureDisp, IOlePicturePersistence
        {
            public ulong Maximum = 4;
            public int SizeError, SaveCalls;
            public bool ChangeDirty, ClearDirtyRequested;
            public int Handle { get { throw new AssertFailedException("A content digest must never inspect a GDI handle."); } }
            public short Type { get { return 1; } }
            public int Width { get { return 100; } }
            public int Height { get { return 100; } }
            public int GetClassID(out Guid id) { id = Guid.Empty; return 0; }
            public int IsDirty() { return ChangeDirty && SaveCalls > 0 ? 0 : 1; }
            public int Load(IStream stream) { throw new NotSupportedException(); }
            public int Save(IStream stream, bool clearDirty)
            {
                SaveCalls++; ClearDirtyRequested = clearDirty;
                stream.Write(new byte[] { 1, 2, 3, 4 }, 4, IntPtr.Zero); return 0;
            }
            public int GetSizeMax(out ulong value) { value = Maximum; return SizeError; }
        }

        private static void OnSta(Action action)
        {
            Exception failure = null;
            var thread = new Thread(() => { try { action(); } catch (Exception error) { failure = error; } });
            thread.SetApartmentState(ApartmentState.STA); thread.Start();
            Assert.IsTrue(thread.Join(30000), "Owned OLE content test did not settle; no host application is involved.");
            if (failure != null) System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(failure).Throw();
        }

        private static object PersistenceRoundtrip(object source)
        {
            IStream stream = null; object result = null;
            try
            {
                var persistence = (IOlePicturePersistence)source;
                Marshal.ThrowExceptionForHR(persistence.GetSizeMax(out ulong maximum));
                Assert.IsTrue(maximum > 0 && maximum <= (ulong)OlePictureLoader.MaximumPersistenceBytes);
                CreateStreamOnHGlobal(IntPtr.Zero, true, out stream);
                Marshal.ThrowExceptionForHR(persistence.Save(stream, false));
                stream.Stat(out System.Runtime.InteropServices.ComTypes.STATSTG state, 1);
                Assert.IsTrue(state.cbSize > 0 && state.cbSize <= (long)maximum);
                stream.Seek(0, 0, IntPtr.Zero);
                Guid iid = typeof(IOlePictureDisp).GUID;
                OleCreatePictureIndirect(IntPtr.Zero, ref iid, true, out result);
                Marshal.ThrowExceptionForHR(((IOlePicturePersistence)result).Load(stream));
                object owned = result; result = null; return owned;
            }
            finally
            {
                if (result != null) Marshal.FinalReleaseComObject(result);
                if (stream != null) Marshal.ReleaseComObject(stream);
            }
        }

        private static byte[] ReadBorrowedMetafile(int handle)
        {
            IntPtr borrowed = new IntPtr(handle);
            uint length = GetMetaFileBitsEx(borrowed, 0, null);
            Assert.IsTrue(length > 0 && length <= OlePictureLoader.MaximumPersistenceBytes, "Borrowed WMF resource is invalid or exceeds the content budget.");
            var bytes = new byte[length];
            Assert.AreEqual(length, GetMetaFileBitsEx(borrowed, length, bytes), "Borrowed WMF read was incomplete.");
            return bytes; // Never delete this handle; its owning picture remains alive.
        }

        private static void WriteOwnedPicture(string path, string format, Color color)
        {
            if (format == "wmf" || format == "emf") { WriteOwnedMetafile(path, format == "emf", color); return; }
            using (var bitmap = new Bitmap(32, 32))
            {
                using (var graphics = Graphics.FromImage(bitmap)) graphics.Clear(color);
                if (format == "ico")
                {
                    IntPtr handle = bitmap.GetHicon();
                    try { using (var icon = Icon.FromHandle(handle)) using (var output = File.Create(path)) icon.Save(output); }
                    finally { Assert.IsTrue(DestroyIcon(handle), "Owned synthetic icon release failed."); }
                }
                else bitmap.Save(path, format == "bmp" ? ImageFormat.Bmp : format == "gif" ? ImageFormat.Gif : ImageFormat.Jpeg);
            }
        }

        [StructLayout(LayoutKind.Sequential)] private struct RectangleBounds { public int Left, Top, Right, Bottom; }
        private static void WriteOwnedMetafile(string path, bool enhanced, Color color)
        {
            var bounds = new RectangleBounds { Right = 847, Bottom = 847 };
            IntPtr dc = enhanced ? CreateEnhMetaFileW(IntPtr.Zero, null, ref bounds, null) : CreateMetaFileW(null);
            Assert.AreNotEqual(IntPtr.Zero, dc, "Owned metafile DC creation failed.");
            IntPtr metafile = IntPtr.Zero, brush = IntPtr.Zero, previous = IntPtr.Zero;
            try
            {
                if (!enhanced) { Assert.AreNotEqual(0, SetMapMode(dc, 8)); Assert.IsTrue(SetWindowExtEx(dc, 32, 32, IntPtr.Zero)); }
                brush = CreateSolidBrush((uint)(color.R | color.G << 8 | color.B << 16));
                Assert.AreNotEqual(IntPtr.Zero, brush);
                previous = SelectObject(dc, brush); Assert.AreNotEqual(IntPtr.Zero, previous);
                Assert.IsTrue(Rectangle(dc, 0, 0, 32, 32));
                Assert.AreNotEqual(IntPtr.Zero, SelectObject(dc, previous)); previous = IntPtr.Zero;
                metafile = enhanced ? CloseEnhMetaFile(dc) : CloseMetaFile(dc); dc = IntPtr.Zero;
                Assert.AreNotEqual(IntPtr.Zero, metafile);
                uint length = enhanced ? GetEnhMetaFileBits(metafile, 0, null) : GetMetaFileBitsEx(metafile, 0, null);
                Assert.IsTrue(length > 0 && length < 1024 * 1024);
                var bytes = new byte[length];
                Assert.AreEqual(length, enhanced ? GetEnhMetaFileBits(metafile, length, bytes) : GetMetaFileBitsEx(metafile, length, bytes));
                using (var output = File.Create(path)) using (var writer = new BinaryWriter(output))
                {
                    if (!enhanced)
                    {
                        // Standard placeable WMF header establishes physical bounds for OleLoadPictureFile.
                        var words = new ushort[] { 0xCDD7, 0x9AC6, 0, 0, 0, 32, 32, 96, 0, 0 };
                        ushort checksum = 0; foreach (ushort word in words) { writer.Write(word); checksum ^= word; }
                        writer.Write(checksum);
                    }
                    writer.Write(bytes);
                }
            }
            finally
            {
                if (dc != IntPtr.Zero)
                {
                    if (previous != IntPtr.Zero) SelectObject(dc, previous);
                    IntPtr aborted = enhanced ? CloseEnhMetaFile(dc) : CloseMetaFile(dc);
                    if (aborted != IntPtr.Zero) { if (enhanced) DeleteEnhMetaFile(aborted); else DeleteMetaFile(aborted); }
                }
                if (brush != IntPtr.Zero) Assert.IsTrue(DeleteObject(brush), "Owned brush release failed.");
                if (metafile != IntPtr.Zero) Assert.IsTrue(enhanced ? DeleteEnhMetaFile(metafile) : DeleteMetaFile(metafile), "Owned metafile release failed.");
            }
        }

        [DllImport("ole32.dll", PreserveSig = false)] private static extern void CreateStreamOnHGlobal(IntPtr memory, [MarshalAs(UnmanagedType.Bool)] bool free, out IStream stream);
        [DllImport("oleaut32.dll", PreserveSig = false)] private static extern void OleCreatePictureIndirect(IntPtr description, ref Guid iid, [MarshalAs(UnmanagedType.Bool)] bool owns, [MarshalAs(UnmanagedType.Interface)] out object picture);
        [DllImport("user32.dll")] private static extern bool DestroyIcon(IntPtr handle);
        [DllImport("gdi32.dll", CharSet = CharSet.Unicode)] private static extern IntPtr CreateMetaFileW(string file);
        [DllImport("gdi32.dll", CharSet = CharSet.Unicode)] private static extern IntPtr CreateEnhMetaFileW(IntPtr reference, string file, ref RectangleBounds bounds, string description);
        [DllImport("gdi32.dll")] private static extern IntPtr CloseMetaFile(IntPtr dc);
        [DllImport("gdi32.dll")] private static extern IntPtr CloseEnhMetaFile(IntPtr dc);
        [DllImport("gdi32.dll")] private static extern bool DeleteMetaFile(IntPtr file);
        [DllImport("gdi32.dll")] private static extern bool DeleteEnhMetaFile(IntPtr file);
        [DllImport("gdi32.dll")] private static extern uint GetMetaFileBitsEx(IntPtr file, uint count, [Out] byte[] bytes);
        [DllImport("gdi32.dll")] private static extern uint GetEnhMetaFileBits(IntPtr file, uint count, [Out] byte[] bytes);
        [DllImport("gdi32.dll")] private static extern int SetMapMode(IntPtr dc, int mode);
        [DllImport("gdi32.dll")] private static extern bool SetWindowExtEx(IntPtr dc, int x, int y, IntPtr previous);
        [DllImport("gdi32.dll")] private static extern IntPtr CreateSolidBrush(uint color);
        [DllImport("gdi32.dll")] private static extern IntPtr SelectObject(IntPtr dc, IntPtr handle);
        [DllImport("gdi32.dll")] private static extern bool Rectangle(IntPtr dc, int left, int top, int right, int bottom);
        [DllImport("gdi32.dll")] private static extern bool DeleteObject(IntPtr handle);
        [TestMethod]
        public void NativeOleImageLoadAndInvalidNativeResultsMatrix()
        {
            foreach (string path in new[] { null, " ", "relative.bmp" }) Assert.ThrowsException<ArgumentException>(() => OlePictureLoader.Load(path));
            string file = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ole-picture-" + Guid.NewGuid().ToString("N") + ".bmp");
            Assert.ThrowsException<FileNotFoundException>(() => OlePictureLoader.Load(file));
            var previous = OlePictureLoader.ReadPicture;
            try
            {
                using (var bitmap = new Bitmap(3, 2)) bitmap.Save(file, ImageFormat.Bmp);
                object picture = OlePictureLoader.Load(file);
                try { Assert.IsTrue(Marshal.IsComObject(picture)); Assert.IsFalse(string.IsNullOrWhiteSpace(OlePictureLoader.Fingerprint(picture))); }
                finally { Marshal.FinalReleaseComObject(picture); }
                OlePictureLoader.ReadPicture = (object path, out object result) => result = null;
                Assert.ThrowsException<InvalidOperationException>(() => OlePictureLoader.Load(file));
                OlePictureLoader.ReadPicture = (object path, out object result) => result = new object();
                Assert.ThrowsException<InvalidOperationException>(() => OlePictureLoader.Load(file));
                Assert.IsNull(OlePictureLoader.Fingerprint(null));
                Assert.ThrowsException<InvalidOperationException>(() => OlePictureLoader.Fingerprint(new object()));
            }
            finally { OlePictureLoader.ReadPicture = previous; File.Delete(file); }
        }
    }
}
