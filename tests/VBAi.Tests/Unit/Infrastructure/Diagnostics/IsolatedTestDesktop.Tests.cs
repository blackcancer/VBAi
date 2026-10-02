using System;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class IsolatedTestDesktopTests
    {
        [TestMethod]
        public void EmptyPrivateDesktopInventoryObservesItsSentinelWithoutSwitchingInput()
        {
            string input = IsolatedTestDesktop.InputDesktopName();
            string desktop = "VBAiTests_" + Guid.NewGuid().ToString("N");
            using (IsolatedTestDesktop.Create(desktop)) Assert.IsFalse(IsolatedTestDesktop.HasWindows(desktop));
            Assert.AreEqual(input, IsolatedTestDesktop.InputDesktopName());
        }
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
        public void OnlyGeneratedTestDesktopsAndBoundedNonNullArgumentsAreAccepted()
        {
            IsolatedTestDesktop.RequireName("VBAiTests_" + Guid.NewGuid().ToString("N"));
            foreach (string invalid in new[] { null, "", "Default", "Winlogon", "WinSta0\\Default", "VBAiTests_invalid", "VBAiTests_" + Guid.NewGuid().ToString("D") })
                Assert.ThrowsException<ArgumentException>(() => IsolatedTestDesktop.RequireName(invalid));
            Assert.ThrowsException<ArgumentException>(() => IsolatedTestDesktop.Quote(null));
            Assert.ThrowsException<ArgumentException>(() => IsolatedTestDesktop.Quote("bad\0argument"));
            Assert.ThrowsException<ArgumentNullException>(() => IsolatedTestDesktop.CommandLine("owned.exe", null));
            Assert.ThrowsException<ArgumentException>(() => IsolatedTestDesktop.CommandLine("owned.exe", new[] { new string('x', 32767) }));
            Assert.AreEqual("\"C:\\Office\\EXCEL.EXE\" /x /automation \"E:\\Owned folder\\seed.xlsx\"",
                IsolatedTestDesktop.CommandLine("C:\\Office\\EXCEL.EXE", new[] { "/x", "/automation", "E:\\Owned folder\\seed.xlsx" }));
        }
    }
}
