using System;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class IsolatedTestDesktopTests
    {
        [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
        private static extern IntPtr CommandLineToArgvW(string command, out int count);
        [DllImport("kernel32.dll")] private static extern IntPtr LocalFree(IntPtr memory);

        [TestMethod]
        public void ArgumentsRoundTripUnicodeEmptyQuotesAndTrailingSlashesWithoutShellEvaluation()
        {
            string[] arguments = { "", "plain", "space here", "E:\\Développement\\test\\", "a\"b",
                "before\\\\\"after", "`$()&|;", "line\r\nbreak", "\\", "\"", "\\\"" };
            string command = IsolatedTestDesktop.CommandLine("C:\\Program Files\\owned.exe", arguments);
            int count; IntPtr pointer = CommandLineToArgvW(command, out count);
            Assert.AreNotEqual(IntPtr.Zero, pointer);
            try
            {
                Assert.AreEqual(arguments.Length + 1, count);
                Assert.AreEqual("C:\\Program Files\\owned.exe", Marshal.PtrToStringUni(Marshal.ReadIntPtr(pointer)));
                for (int index = 0; index < arguments.Length; index++)
                    Assert.AreEqual(arguments[index], Marshal.PtrToStringUni(Marshal.ReadIntPtr(pointer, (index + 1) * IntPtr.Size)));
            }
            finally { LocalFree(pointer); }
        }

        [TestMethod]
        public void AnInactiveExactDesktopIsRequiredWithoutFallbackToTheInputDesktop()
        {
            string expected = "VBAiTests_" + Guid.NewGuid().ToString("N");
            IsolatedTestDesktop.RequireObserved(expected, expected.ToUpperInvariant(), "Default");
            Assert.ThrowsException<InvalidOperationException>(() => IsolatedTestDesktop.RequireObserved(expected, "Default", "Default"));
            Assert.ThrowsException<InvalidOperationException>(() => IsolatedTestDesktop.RequireObserved(expected, expected, expected));
            Assert.ThrowsException<InvalidOperationException>(() => IsolatedTestDesktop.RequireObserved(expected, null, "Default"));
            Assert.ThrowsException<InvalidOperationException>(() => IsolatedTestDesktop.RequireObserved(expected, expected, null));
        }

        [TestMethod]
        public void EmptyValidatedDesktopIsDistinctFromAnErroredOrPartialInventory()
        {
            Assert.IsFalse(IsolatedTestDesktop.RequireWindowInventory(false, 0, 0));
            Assert.IsFalse(IsolatedTestDesktop.RequireWindowInventory(true, 0, 0));
            Assert.IsTrue(IsolatedTestDesktop.RequireWindowInventory(true, 1, 0));
            Assert.IsTrue(IsolatedTestDesktop.RequireWindowInventory(true, 8192, 0));
            foreach (int count in new[] { -1, 1, 8192, 8193 })
                Assert.ThrowsException<InvalidOperationException>(() => IsolatedTestDesktop.RequireWindowInventory(false, count, 0));
            Assert.ThrowsException<InvalidOperationException>(() => IsolatedTestDesktop.RequireWindowInventory(false, 0, 5));
            Assert.ThrowsException<InvalidOperationException>(() => IsolatedTestDesktop.RequireWindowInventory(true, 8193, 0));
        }

        [TestMethod]
        public void OnlyGeneratedTestDesktopsAndBoundedNonNullArgumentsAreAccepted()
        {
            IsolatedTestDesktop.RequireName("VBAiTests_" + Guid.NewGuid().ToString("N"));
            foreach (string invalid in new[] { null, "", "Default", "Winlogon", "WinSta0\\Default", "VBAiTests_invalid", "VBAiTests_" + Guid.NewGuid().ToString("D") })
                Assert.ThrowsException<ArgumentException>(() => IsolatedTestDesktop.RequireName(invalid));
            Assert.ThrowsException<ArgumentException>(() => IsolatedTestDesktop.Quote(null));
            Assert.ThrowsException<ArgumentException>(() => IsolatedTestDesktop.Quote("bad\0argument"));
            Assert.ThrowsException<ArgumentNullException>(() => IsolatedTestDesktop.CommandLine("owned.exe", null));
            Assert.ThrowsException<ArgumentException>(() => IsolatedTestDesktop.CommandLine("owned.exe", new[] { new string('x', 32767) }));
        }
    }
}
