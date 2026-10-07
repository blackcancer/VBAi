using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    /// <summary>Qualifies retained-handle image reads without starting a host or using Process.MainModule.</summary>
    [TestClass, TestCategory("Unit")]
    public sealed class ExcelOwnedProcessImageTests
    {
        [DllImport("kernel32.dll", EntryPoint = "GetModuleFileNameW", CharSet = CharSet.Unicode, ExactSpelling = true, SetLastError = true)]
        private static extern uint GetCurrentImage(IntPtr module, StringBuilder path, int size);

        [TestMethod]
        public void NativeHandleReadMatchesCurrentTesthostImageWithoutModuleEnumeration()
        {
            var expected = new StringBuilder(ExcelOwnedProcessImage.MaximumPathCharacters);
            uint size = GetCurrentImage(IntPtr.Zero, expected, expected.Capacity);
            Assert.IsTrue(size > 0 && size < expected.Capacity, "Current testhost module identity must be available independently.");
            using (var process = Process.GetCurrentProcess())
            {
                string actual = ExcelOwnedProcessImage.Read(process.Handle);
                Assert.AreEqual(Path.GetFullPath(expected.ToString()), actual, true);
            }
        }

        [TestMethod]
        public void NativeFailurePreservesExactWin32ErrorAfterOneBoundedQuery()
        {
            int calls = 0;
            ExcelOwnedProcessImage.ImageQuery query = (IntPtr handle, int flags, StringBuilder path, ref int size) =>
            {
                calls++;
                Assert.AreEqual(new IntPtr(42), handle);
                Assert.AreEqual(0, flags, "Win32 paths are required, not device paths.");
                Assert.AreEqual(ExcelOwnedProcessImage.MaximumPathCharacters, size);
                return false;
            };
            var failure = Assert.ThrowsException<Win32Exception>(() => ExcelOwnedProcessImage.Read(new IntPtr(42), query, () => 5));
            Assert.AreEqual(5, failure.NativeErrorCode);
            Assert.AreEqual(1, calls, "A failed identity read must not retry or authorize another process launch.");
        }

        [TestMethod]
        public void OnlyCompleteNativePathCanSatisfyExactPidPathAndStartGuards()
        {
            string image = Path.Combine(Path.GetTempPath(), "EXCEL.EXE");
            ExcelOwnedProcessImage.ImageQuery query = (IntPtr handle, int flags, StringBuilder path, ref int size) =>
            {
                path.Append(image);
                size = image.Length;
                return true;
            };
            string actual = ExcelOwnedProcessImage.Read(new IntPtr(42), query, () => { Assert.Fail("Success must not consume a stale last error."); return 0; });
            const string started = "2026-09-30T21:22:20.6505688Z";
            ExcelOwnedBootstrapPlan.VerifyAttachedIdentity(42, image, started, 42, actual, started);
            Assert.ThrowsException<InvalidOperationException>(() => ExcelOwnedBootstrapPlan.VerifyAttachedIdentity(42, Path.Combine(Path.GetTempPath(), "other", "EXCEL.EXE"), started, 42, actual, started));
            Assert.ThrowsException<InvalidOperationException>(() => ExcelOwnedBootstrapPlan.VerifyAttachedIdentity(42, image, started, 43, actual, started));
            Assert.ThrowsException<InvalidOperationException>(() => ExcelOwnedBootstrapPlan.VerifyAttachedIdentity(42, image, started, 42, actual, started + "1"));
        }

        [DataTestMethod]
        [DataRow(0)]
        [DataRow(32768)]
        [DataRow(12)]
        public void EmptyTruncatedOrInconsistentNativeLengthCannotProveIdentity(int returnedSize)
        {
            ExcelOwnedProcessImage.ImageQuery query = (IntPtr handle, int flags, StringBuilder path, ref int size) =>
            {
                path.Append("C:\\x.exe");
                size = returnedSize;
                return true;
            };
            Assert.ThrowsException<InvalidOperationException>(() => ExcelOwnedProcessImage.Read(new IntPtr(42), query, () => 0));
        }

        [TestMethod]
        public void InvalidHandlesAreRejectedBeforeNativeQuery()
        {
            int calls = 0;
            ExcelOwnedProcessImage.ImageQuery query = (IntPtr handle, int flags, StringBuilder path, ref int size) => { calls++; return false; };
            Assert.ThrowsException<ArgumentException>(() => ExcelOwnedProcessImage.Read(IntPtr.Zero, query, () => 6));
            Assert.ThrowsException<ArgumentException>(() => ExcelOwnedProcessImage.Read(new IntPtr(-1), query, () => 6));
            Assert.AreEqual(0, calls);
        }
    }
}
