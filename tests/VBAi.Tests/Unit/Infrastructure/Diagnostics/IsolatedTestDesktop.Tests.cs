using System;
using System.ComponentModel;
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

        [TestMethod]
        public void ThreadDesktopFailuresRetainTheApiThreadAndZeroErrorWithoutClaimingSuccess()
        {
            Assert.ThrowsException<ArgumentNullException>(() => IsolatedTestDesktop.RequireThreadDesktopName(null));
            foreach (int error in new[] { 0, 5, 87 })
            {
                var observation = new IsolatedTestDesktop.ThreadDesktopObservation(41, 0, null, error, 0);
                var failure = Assert.ThrowsException<Win32Exception>(() => IsolatedTestDesktop.RequireThreadDesktopName(observation));
                Assert.AreEqual(error, failure.NativeErrorCode);
                StringAssert.Contains(failure.Message, "GetThreadDesktop returned NULL for thread 41");
                StringAssert.Contains(failure.Message, "Desktop identity is unproved");
                Assert.AreEqual(0L, observation.Handle);
                Assert.AreEqual(error, observation.DesktopError);
            }
            foreach (int error in new[] { 0, 5, 122 })
            {
                var observation = new IsolatedTestDesktop.ThreadDesktopObservation(42, 128, null, 0, error);
                var failure = Assert.ThrowsException<Win32Exception>(() => IsolatedTestDesktop.RequireThreadDesktopName(observation));
                Assert.AreEqual(error, failure.NativeErrorCode);
                StringAssert.Contains(failure.Message, "desktop name is unavailable for thread 42");
                Assert.AreEqual(128L, observation.Handle);
                Assert.AreEqual(error, observation.NameError);
            }
            Assert.AreEqual("Default", IsolatedTestDesktop.RequireThreadDesktopName(
                new IsolatedTestDesktop.ThreadDesktopObservation(43, 256, "Default", 0, 0)));
            Assert.ThrowsException<Win32Exception>(() => IsolatedTestDesktop.RequireThreadDesktopName(
                new IsolatedTestDesktop.ThreadDesktopObservation(43, 256, "Default", 0, 5)));
            Assert.ThrowsException<Win32Exception>(() => IsolatedTestDesktop.RequireThreadDesktopName(
                new IsolatedTestDesktop.ThreadDesktopObservation(43, 256, "", 0, 0)));
        }

        [TestMethod]
        public void NamedDesktopInventoryHandlesEmptyCompletePartialAndOverBoundEnumerationsAsOneBank()
        {
            string desktop = "VBAiTests_" + Guid.NewGuid().ToString("N");
            foreach (int count in new[] { 0, 1, 8192 })
            {
                int names = 0, visits = 0;
                IsolatedTestDesktop.InventoryNamedWindows(desktop, () => { names++; return desktop; }, callback => {
                    for (int index = 0; index < count; index++) Assert.IsTrue(callback(new IntPtr(index + 1)));
                    return new IsolatedTestDesktop.WindowEnumeration { Completed = count != 0, Error = 0 };
                }, unused => { visits++; return true; });
                Assert.AreEqual(2, names); Assert.AreEqual(count, visits);
            }
            foreach (int failure in new[] { 0, 1, 2, 3 })
            {
                int names = 0, visits = 0;
                Assert.ThrowsException<InvalidOperationException>(() => IsolatedTestDesktop.InventoryNamedWindows(desktop,
                    () => { names++; return desktop; }, callback => {
                        if (failure == 0) return null;
                        if (failure == 1) return new IsolatedTestDesktop.WindowEnumeration { Completed = false, Error = 5 };
                        int bound = failure == 2 ? 1 : 8193;
                        for (int index = 0; index < bound; index++) if (!callback(new IntPtr(index + 1))) break;
                        return new IsolatedTestDesktop.WindowEnumeration { Completed = false, Error = 0 };
                    }, unused => { visits++; return true; }));
                Assert.AreEqual(2, names); Assert.AreEqual(failure == 3 ? 8192 : failure == 2 ? 1 : 0, visits);
            }
            Assert.ThrowsException<InvalidOperationException>(() => IsolatedTestDesktop.InventoryNamedWindows(desktop,
                () => desktop, callback => { Assert.IsFalse(callback(new IntPtr(1))); return new IsolatedTestDesktop.WindowEnumeration(); }, unused => false));
        }

        [TestMethod]
        public void NamedDesktopInventoryPreservesVisitorNativeAndIdentityRecheckFailures()
        {
            string desktop = "VBAiTests_" + Guid.NewGuid().ToString("N");
            var original = new InvalidOperationException("Original visitor failure");
            var recheck = new InvalidOperationException("Desktop name recheck failed");
            foreach (int failure in new[] { 0, 1, 2, 3 })
            {
                int names = 0;
                Action run = () => IsolatedTestDesktop.InventoryNamedWindows(desktop,
                    () => { if (++names == 1 || failure == 0) return desktop; if (failure == 1) return "Default"; throw recheck; },
                    callback => {
                        Assert.IsFalse(callback(new IntPtr(1)));
                        if (failure == 3) throw new InvalidOperationException("Enumeration also failed");
                        return new IsolatedTestDesktop.WindowEnumeration { Completed = false };
                    }, unused => throw original);
                if (failure == 0) Assert.AreSame(original, Assert.ThrowsException<InvalidOperationException>(run));
                else
                {
                    var aggregate = Assert.ThrowsException<AggregateException>(run).Flatten();
                    CollectionAssert.Contains(aggregate.InnerExceptions, original);
                    Assert.AreEqual(failure == 3 ? 3 : 2, aggregate.InnerExceptions.Count);
                    if (failure >= 2) CollectionAssert.Contains(aggregate.InnerExceptions, recheck);
                }
                Assert.AreEqual(2, names);
            }
            Assert.AreSame(original, Assert.ThrowsException<InvalidOperationException>(() => IsolatedTestDesktop.InventoryNamedWindows(
                desktop, () => desktop, unused => throw original, unused => { Assert.Fail(); return true; })));
            int nameReads = 0;
            Assert.ThrowsException<InvalidOperationException>(() => IsolatedTestDesktop.InventoryNamedWindows(desktop,
                () => { nameReads++; return "Default"; }, unused => { Assert.Fail(); return null; }, unused => { Assert.Fail(); return true; }));
            Assert.AreEqual(1, nameReads);
            Assert.ThrowsException<ArgumentNullException>(() => IsolatedTestDesktop.InventoryNamedWindows(desktop, null, null, null));
            Assert.ThrowsException<ArgumentException>(() => IsolatedTestDesktop.InventoryNamedWindows("Default",
                () => { Assert.Fail(); return null; }, unused => null, unused => true));
        }
    }
}
