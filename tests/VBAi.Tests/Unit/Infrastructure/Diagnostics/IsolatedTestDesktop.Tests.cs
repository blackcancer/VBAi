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

        [TestMethod]
        public void ForeignOfficeWindowAcceptanceUsesExactSentinelAndTargetMembershipWithoutForeignThreadQueries()
        {
            string desktop = "VBAiTests_" + Guid.NewGuid().ToString("N");
            var sentinel = new IsolatedTestDesktop.WindowIdentity(new IntPtr(11), 100, 200);
            var office = new IsolatedTestDesktop.WindowIdentity(new IntPtr(12), 101, 201);
            var privateInventory = new IsolatedTestDesktop.WindowInventory(desktop, true, sentinel, office);
            var inputInventory = new IsolatedTestDesktop.WindowInventory("Default", true,
                new IsolatedTestDesktop.WindowIdentity(new IntPtr(13), 102, 202));
            IsolatedTestDesktop.ValidateOfficeWindowInventories(desktop, 101, sentinel, privateInventory, inputInventory, true, office.Window);
            foreach (var wrongSentinel in new[] {
                new IsolatedTestDesktop.WindowIdentity(new IntPtr(99), 100, 200),
                new IsolatedTestDesktop.WindowIdentity(sentinel.Window, 999, 200),
                new IsolatedTestDesktop.WindowIdentity(sentinel.Window, 100, 999) })
                Assert.ThrowsException<InvalidOperationException>(() => IsolatedTestDesktop.ValidateOfficeWindowInventories(
                    desktop, 101, wrongSentinel, privateInventory, inputInventory, true, office.Window));
            Assert.ThrowsException<InvalidOperationException>(() => IsolatedTestDesktop.ValidateOfficeWindowInventories(
                desktop, 101, sentinel, privateInventory, inputInventory, true, new IntPtr(99)));
        }

        [TestMethod]
        public void OfficeInventoryRejectsInputLeakageIncompleteEnumerationAndInputDesktopSwitch()
        {
            string desktop = "VBAiTests_" + Guid.NewGuid().ToString("N");
            var sentinel = new IsolatedTestDesktop.WindowIdentity(new IntPtr(11), 100, 200);
            var office = new IsolatedTestDesktop.WindowIdentity(new IntPtr(12), 101, 201);
            var privateInventory = new IsolatedTestDesktop.WindowInventory(desktop, true, sentinel, office);
            foreach (var badInput in new[] {
                new IsolatedTestDesktop.WindowInventory("Default", true, office),
                new IsolatedTestDesktop.WindowInventory("Default", false),
                new IsolatedTestDesktop.WindowInventory(desktop, true),
                new IsolatedTestDesktop.WindowInventory("Default", true, new IsolatedTestDesktop.WindowIdentity(new IntPtr(99), 0, 0)) })
                Assert.ThrowsException<InvalidOperationException>(() => IsolatedTestDesktop.ValidateOfficeWindowInventories(
                    desktop, 101, sentinel, privateInventory, badInput, true, office.Window));
            var validInput = new IsolatedTestDesktop.WindowInventory("Default", true);
            Assert.ThrowsException<InvalidOperationException>(() => IsolatedTestDesktop.ValidateOfficeWindowInventories(
                desktop, 101, sentinel, new IsolatedTestDesktop.WindowInventory(desktop, false, sentinel, office), validInput, true, office.Window));
            Assert.ThrowsException<InvalidOperationException>(() => IsolatedTestDesktop.ValidateOfficeWindowInventories(
                desktop, 101, sentinel, new IsolatedTestDesktop.WindowInventory("Default", true, sentinel, office), validInput, true, office.Window));
        }

        [TestMethod]
        public void EarlyOfficeDiscoveryAllowsNoHostWindowOnlyWithLiveSentinelAndCompleteInactiveInventories()
        {
            string desktop = "VBAiTests_" + Guid.NewGuid().ToString("N");
            var sentinel = new IsolatedTestDesktop.WindowIdentity(new IntPtr(11), 100, 200);
            var privateInventory = new IsolatedTestDesktop.WindowInventory(desktop, true, sentinel);
            var inputInventory = new IsolatedTestDesktop.WindowInventory("Default", true);
            IsolatedTestDesktop.ValidateOfficeWindowInventories(desktop, 101, sentinel, privateInventory, inputInventory, false, IntPtr.Zero);
            Assert.ThrowsException<InvalidOperationException>(() => IsolatedTestDesktop.ValidateOfficeWindowInventories(
                desktop, 101, sentinel, privateInventory, inputInventory, true, IntPtr.Zero));
            Assert.ThrowsException<InvalidOperationException>(() => IsolatedTestDesktop.ValidateOfficeWindowInventories(
                desktop, 101, sentinel, new IsolatedTestDesktop.WindowInventory(desktop, true), inputInventory, false, IntPtr.Zero));
            Assert.ThrowsException<InvalidOperationException>(() => IsolatedTestDesktop.ValidateOfficeWindowInventories(
                desktop, 100, sentinel, privateInventory, inputInventory, false, IntPtr.Zero));
        }
    }
}
