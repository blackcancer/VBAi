using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class NativeWindowEnumerationOwnershipTests
    {
        [DllImport("user32.dll", SetLastError = true)] private static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId);
        [DllImport("user32.dll")] private static extern bool IsWindow(IntPtr window);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern IntPtr CreateWindowExW(uint extended, string cls, string title, uint style,
            int x, int y, int width, int height, IntPtr parent, IntPtr menu, IntPtr instance, IntPtr data);
        [DllImport("user32.dll")] private static extern bool DestroyWindow(IntPtr window);

        [STATestMethod]
        public void ActualDestroyedOwnedHandleIsConfirmedGoneWithoutInputOrReplacementWindow()
        {
            IntPtr window = CreateWindowExW(0, "STATIC", "", 0, 0, 0, 1, 1, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero);
            Assert.AreNotEqual(IntPtr.Zero, window);
            bool destroyed = false;
            try
            {
                Assert.AreEqual((uint)Process.GetCurrentProcess().Id,
                    NativeWindowEnumerationOwnership.ReadProcessId(window, GetWindowThreadProcessId, Marshal.GetLastWin32Error, IsWindow));
                Assert.IsTrue(DestroyWindow(window)); destroyed = true;
                Assert.IsNull(NativeWindowEnumerationOwnership.ReadProcessId(window, GetWindowThreadProcessId, Marshal.GetLastWin32Error, IsWindow));
            }
            finally { if (!destroyed) DestroyWindow(window); }
        }

        [DataTestMethod, DataRow(0, false), DataRow(5, false), DataRow(1400, true)]
        public void UnknownErrorOrLiveReusedHandleCannotBeIgnored(int error, bool live)
        {
            uint Failed(IntPtr h, out uint pid) { pid = 0; return 0; }
            Assert.ThrowsException<Win32Exception>(() => NativeWindowEnumerationOwnership.ReadProcessId(new IntPtr(42), Failed, () => error, h => live));
        }

        [TestMethod]
        public void ForeignLiveOwnerIsReturnedForTheCallerToFilterWithoutReadingItsClassOrTitle()
        {
            uint Foreign(IntPtr h, out uint pid) { pid = 1234; return 5678; }
            Assert.AreEqual((uint)1234, NativeWindowEnumerationOwnership.ReadProcessId(new IntPtr(42), Foreign,
                () => throw new AssertFailedException("No failed native read."), h => throw new AssertFailedException("No lifetime fallback.")));
        }

        [TestMethod]
        public void ZeroHandleCannotAuthorizeWindowAbsence()
        {
            uint Never(IntPtr h, out uint pid) { throw new AssertFailedException("Invalid handle must fail before native reads."); }
            Assert.ThrowsException<ArgumentException>(() => NativeWindowEnumerationOwnership.ReadProcessId(IntPtr.Zero, Never, () => 1400, h => false));
        }
    }
}
