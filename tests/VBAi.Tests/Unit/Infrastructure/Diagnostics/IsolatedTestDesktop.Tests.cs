using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class IsolatedTestDesktopTests
    {
        [TestMethod]
        public void EmptyPrivateDesktopInventoryDoesNotSwitchInput()
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
        public void OfficeSwitchesRemainLiteralWhileDocumentPathsRemainQuoted()
        {
            Assert.AreEqual("\"C:\\Office\\EXCEL.EXE\" /x /automation \"E:\\Owned folder\\seed.xlsx\"",
                IsolatedTestDesktop.CommandLine("C:\\Office\\EXCEL.EXE", new[] { "/x", "/automation", "E:\\Owned folder\\seed.xlsx" }));
            Assert.AreEqual("\"C:\\Office\\WINWORD.EXE\" /a \"E:\\Owned folder\\seed.docx\"",
                IsolatedTestDesktop.CommandLine("C:\\Office\\WINWORD.EXE", new[] { "/a", "E:\\Owned folder\\seed.docx" }));
            foreach (string argument in new[] { "/a ", "/aOther", "/automationOther", "/xOther", "/a\" /mUntrusted", "" })
                Assert.AreEqual("\"owned.exe\" " + IsolatedTestDesktop.Quote(argument),
                    IsolatedTestDesktop.CommandLine("owned.exe", new[] { argument }));
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
                IsolatedTestDesktop.InventoryNamedWindows(desktop, () => { names++; return desktop; }, callback =>
                {
                    for (int index = 0; index < count; index++) Assert.IsTrue(callback(new IntPtr(index + 1)));
                    return new IsolatedTestDesktop.WindowEnumeration { Completed = count != 0, Error = 0 };
                }, unused => { visits++; return true; });
                Assert.AreEqual(2, names); Assert.AreEqual(count, visits);
            }
            foreach (int failure in new[] { 0, 1, 2, 3 })
            {
                int names = 0, visits = 0;
                Assert.ThrowsException<InvalidOperationException>(() => IsolatedTestDesktop.InventoryNamedWindows(desktop,
                    () => { names++; return desktop; }, callback =>
                    {
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
                    callback =>
                    {
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
        [TestMethod]
        public void SuccessfulDesktopCloseIsClaimedBeforeDeliveryAndRepeatedDisposeHasNoNativeCall()
        {
            int calls = 0; bool claimedBeforeCall = false;
            IsolatedTestDesktop.DesktopLease lease = null;
            lease = new IsolatedTestDesktop.DesktopLease(new IntPtr(123), handle =>
            {
                calls++; claimedBeforeCall = lease.CloseAttempted;
                return new IsolatedTestDesktop.DesktopCloseResult(true, 0);
            });
            lease.Dispose(); lease.Dispose();
            Assert.AreEqual(1, calls);
            Assert.IsTrue(claimedBeforeCall);
            Assert.IsTrue(lease.CloseAttempted);
            Assert.IsTrue(lease.CloseSucceeded);
            Assert.AreEqual(0, lease.CloseError);
            Assert.AreEqual(IntPtr.Zero, lease.Handle);
            Assert.IsNull(lease.CloseFailure);
        }
        [TestMethod]
        public void FailedDesktopClosePreservesOriginalErrorHandleAndFailureWithoutNativeRetry()
        {
            int calls = 0; IntPtr original = new IntPtr(123);
            var lease = new IsolatedTestDesktop.DesktopLease(original, handle =>
            {
                calls++;
                return new IsolatedTestDesktop.DesktopCloseResult(false, 170);
            });
            var first = Assert.ThrowsException<Win32Exception>(() => lease.Dispose());
            var second = Assert.ThrowsException<Win32Exception>(() => lease.Dispose());
            Assert.AreEqual(1, calls);
            Assert.AreEqual(170, first.NativeErrorCode);
            Assert.AreSame(first, second);
            Assert.AreSame(first, lease.CloseFailure);
            Assert.AreEqual(170, lease.CloseError);
            Assert.IsTrue(lease.CloseAttempted);
            Assert.IsFalse(lease.CloseSucceeded);
            Assert.AreEqual(original, lease.Handle);
        }
        [TestMethod]
        public void OwnedShutdownPublishesTerminalOnlyAfterBothClosesAndOriginalChildRelease()
        {
            string order = ""; bool childHeld = true, closeObservedChildHeld = false;
            var lease = new IsolatedTestDesktop.DesktopLease(new IntPtr(123), handle =>
            {
                closeObservedChildHeld = childHeld; order += "desktop;";
                return new IsolatedTestDesktop.DesktopCloseResult(true, 0);
            });
            IsolatedTestDesktop.CompleteOwnedShutdown(() => order += "sentinel;", lease,
                () => { childHeld = false; order += "child;"; }, () => order += "terminal;");
            Assert.IsTrue(closeObservedChildHeld);
            Assert.IsTrue(lease.CloseSucceeded);
            Assert.IsFalse(childHeld);
            Assert.AreEqual("sentinel;desktop;child;terminal;", order);
        }
        [TestMethod]
        public void DesktopCloseFailureAfterObservedChildExitRetainsHandleAndCannotPublishSuccess()
        {
            int closes = 0, childReleases = 0, terminals = 0; bool sentinelClosed = false;
            var lease = new IsolatedTestDesktop.DesktopLease(new IntPtr(123), handle =>
            {
                closes++;
                return new IsolatedTestDesktop.DesktopCloseResult(false, 170);
            });
            // The original child's exit has been observed; its retained handle must survive close failure.
            var failure = Assert.ThrowsException<Win32Exception>(() => IsolatedTestDesktop.CompleteOwnedShutdown(
                () => sentinelClosed = true, lease, () => childReleases++, () => terminals++));
            Assert.IsTrue(sentinelClosed);
            Assert.AreEqual(0, childReleases);
            Assert.AreEqual(0, terminals);
            Exception closeFailure;
            Assert.IsTrue(IsolatedTestDesktop.PrepareRefusal(lease, true, out closeFailure));
            Assert.AreSame(failure, closeFailure);
            // Desktop uncertainty alone must also retain ownership if another path released its child reference.
            Assert.IsTrue(IsolatedTestDesktop.PrepareRefusal(lease, false, out closeFailure));
            Assert.AreSame(failure, closeFailure);
            Assert.AreEqual(1, closes);
            Assert.AreEqual(new IntPtr(123), lease.Handle);
        }
        [TestMethod]
        public void FailedSentinelClosePreventsCreatorLeaseCloseChildReleaseAndTerminalPublication()
        {
            int closes = 0, childReleases = 0, terminals = 0;
            var lease = new IsolatedTestDesktop.DesktopLease(new IntPtr(123), handle =>
            {
                closes++;
                return new IsolatedTestDesktop.DesktopCloseResult(true, 0);
            });
            Assert.ThrowsException<InvalidOperationException>(() => IsolatedTestDesktop.CompleteOwnedShutdown(
                () => { throw new InvalidOperationException("Original sentinel close failed."); }, lease,
                () => childReleases++, () => terminals++));
            Exception closeFailure;
            Assert.IsTrue(IsolatedTestDesktop.PrepareRefusal(lease, true, out closeFailure));
            Assert.IsNull(closeFailure);
            Assert.IsFalse(lease.CloseAttempted);
            Assert.AreEqual(0, closes);
            Assert.AreEqual(0, childReleases);
            Assert.AreEqual(0, terminals);
            Assert.AreEqual(new IntPtr(123), lease.Handle);
        }
        [TestMethod]
        public void EarlyRefusalWithoutChildOrSentinelClosesUnattemptedLeaseExactlyOnce()
        {
            int closes = 0;
            var lease = new IsolatedTestDesktop.DesktopLease(new IntPtr(123), handle =>
            {
                closes++;
                return new IsolatedTestDesktop.DesktopCloseResult(true, 0);
            });
            Exception closeFailure;
            Assert.IsFalse(IsolatedTestDesktop.PrepareRefusal(lease, false, out closeFailure));
            Assert.IsNull(closeFailure);
            Assert.IsFalse(IsolatedTestDesktop.PrepareRefusal(lease, false, out closeFailure));
            Assert.AreEqual(1, closes);
            Assert.IsTrue(lease.CloseSucceeded);
        }
        [TestMethod]
        public void FailedEarlyRefusalCloseKeepsDesktopUncertaintyAndNeverReplaysDelivery()
        {
            int closes = 0;
            var lease = new IsolatedTestDesktop.DesktopLease(new IntPtr(123), handle =>
            {
                closes++;
                return new IsolatedTestDesktop.DesktopCloseResult(false, 6);
            });
            Exception first, second;
            Assert.IsTrue(IsolatedTestDesktop.PrepareRefusal(lease, false, out first));
            Assert.IsTrue(IsolatedTestDesktop.PrepareRefusal(lease, false, out second));
            Assert.AreSame(first, second);
            Assert.AreEqual(6, ((Win32Exception)first).NativeErrorCode);
            Assert.AreEqual(1, closes);
            Assert.AreEqual(new IntPtr(123), lease.Handle);
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
        [TestMethod]
        public void FailedNativeInventoriesRemainRefusedAndExposeDistinctPrivateInputDiagnostics()
        {
            string desktop = "VBAiTests_" + Guid.NewGuid().ToString("N");
            var sentinel = new IsolatedTestDesktop.WindowIdentity(new IntPtr(11), 100, 200);
            var validPrivate = new IsolatedTestDesktop.WindowInventory(desktop, true, sentinel);
            var validInput = new IsolatedTestDesktop.WindowInventory("Default", true);
            var nativeFailure = new IsolatedTestDesktop.WindowInventory(desktop, false, 5, true, "None",
                IntPtr.Zero, 0, 0, 0, new[] { sentinel });
            var callbackFailure = new IsolatedTestDesktop.WindowInventory("Default", false, 1400, false,
                "InvalidWindowIdentity", new IntPtr(99), 0, 0, 1400,
                new[] { new IsolatedTestDesktop.WindowIdentity(new IntPtr(99), 0, 0) });
            var privateError = Assert.ThrowsException<InvalidOperationException>(() =>
                IsolatedTestDesktop.ValidateOfficeWindowInventories(desktop, 101, sentinel,
                    nativeFailure, validInput, false, IntPtr.Zero));
            StringAssert.Contains(privateError.Message, "Private={Desktop=" + desktop + ",Complete=False,EnumBOOL=False,EnumError=5");
            StringAssert.Contains(privateError.Message, "Input={Desktop=Default,Complete=True");
            var inputError = Assert.ThrowsException<InvalidOperationException>(() =>
                IsolatedTestDesktop.ValidateOfficeWindowInventories(desktop, 101, sentinel,
                    validPrivate, callbackFailure, false, IntPtr.Zero));
            StringAssert.Contains(inputError.Message, "Input={Desktop=Default,Complete=False,EnumBOOL=False,EnumError=1400");
            StringAssert.Contains(inputError.Message, "CallbackStop=InvalidWindowIdentity,FailedHWND=0x63,FailedPID=0,FailedTID=0,IdentityError=1400");
        }
        [TestMethod]
        public void InventoryCapAndDesktopMismatchCannotBecomeAcceptedThroughDiagnosticInstrumentation()
        {
            string desktop = "VBAiTests_" + Guid.NewGuid().ToString("N");
            var sentinel = new IsolatedTestDesktop.WindowIdentity(new IntPtr(11), 100, 200);
            var input = new IsolatedTestDesktop.WindowInventory("Default", true);
            // Even an inconsistent raw successful BOOL cannot override a recorded callback stop.
            var capped = new IsolatedTestDesktop.WindowInventory(desktop, true, 234, true,
                "WindowLimitExceeded", new IntPtr(12), 999, 998, 0, new[] { sentinel });
            Assert.IsFalse(capped.Complete);
            var error = Assert.ThrowsException<InvalidOperationException>(() =>
                IsolatedTestDesktop.ValidateOfficeWindowInventories(desktop, 101, sentinel, capped, input, false, IntPtr.Zero));
            StringAssert.Contains(error.Message, "CallbackStop=WindowLimitExceeded");
            var wrongDesktop = new IsolatedTestDesktop.WindowInventory("Other", true, 0, true, "None",
                IntPtr.Zero, 0, 0, 0, new[] { sentinel });
            var mismatch = Assert.ThrowsException<InvalidOperationException>(() =>
                IsolatedTestDesktop.ValidateOfficeWindowInventories(desktop, 101, sentinel, wrongDesktop, input, false, IntPtr.Zero));
            StringAssert.Contains(mismatch.Message, "Expected=" + desktop);
            StringAssert.Contains(mismatch.Message, "Private={Desktop=Other");
        }
    }
}
