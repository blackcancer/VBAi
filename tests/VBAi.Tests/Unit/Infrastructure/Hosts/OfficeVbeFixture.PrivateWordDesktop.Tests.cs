using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class OfficeVbeFixturePrivateWordDesktopTests
    {
        private const string Executable = @"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE";
        private static readonly string Hash = new string('A', 64);

        [TestMethod]
        public void DisabledOfficeOptInIsInconclusiveBeforeAnyPrivateDesktopObservation()
        {
            string name = "VBAiTests_" + Guid.NewGuid().ToString("N");
            int observations = 0;
            Action<string> observe = unused => observations++;
            foreach (string kind in new[] { "Word", "Access", "PowerPoint", "Publisher" })
                foreach (string optIn in new[] { null, "", "0", "true", "2" })
                    Assert.ThrowsException<AssertInconclusiveException>(() =>
                        OfficeVbeFixture.RequireEnabledOfficeDesktop(kind, optIn, name, null, observe));
            Assert.AreEqual(0, observations,
                "Disabled Office tests must skip before desktop comparison, native inventory or activation.");
        }

        [TestMethod]
        public void EnabledOfficeOptInValidatesPrivateWordDesktopBeforeAnyHostInventory()
        {
            string name = "VBAiTests_" + Guid.NewGuid().ToString("N");
            int observations = 0;
            Action<string> observe = actual => { Assert.AreEqual(name, actual); observations++; };
            Assert.IsNull(OfficeVbeFixture.RequireEnabledOfficeDesktop("Word", "1", null, null, observe));
            Assert.AreEqual(0, observations);
            Assert.AreEqual(name, OfficeVbeFixture.RequireEnabledOfficeDesktop("Word", "1", name, name, observe));
            Assert.AreEqual(1, observations);
            foreach (string configured in new[] { null, "", "wrong" })
                Assert.ThrowsException<InvalidOperationException>(() =>
                    OfficeVbeFixture.RequireEnabledOfficeDesktop("Word", "1", name, configured, observe));
            Assert.ThrowsException<InvalidOperationException>(() =>
                OfficeVbeFixture.RequireEnabledOfficeDesktop("Access", "1", name, name, observe));
            Assert.AreEqual(1, observations, "Pair and host refusals must precede desktop/native access.");
        }

        [TestMethod]
        public void PrivateWordDesktopRequiresWorkerAndTestNamesToMatchBeforeHostActivation()
        {
            string name = "VBAiTests_" + Guid.NewGuid().ToString("N");
            int observations = 0;
            Action<string> observe = actual => { Assert.AreEqual(name, actual); observations++; };
            Assert.IsNull(OfficeVbeFixture.RequirePrivateWordDesktop("Word", null, null, observe));
            Assert.AreEqual(0, observations);
            Assert.AreEqual(name, OfficeVbeFixture.RequirePrivateWordDesktop("Word", name, name, observe));
            Assert.AreEqual(1, observations);
            Assert.AreEqual(name, OfficeVbeFixture.RequirePrivateWordDesktop("Word", null, name, observe));
            Assert.AreEqual(2, observations);
            foreach (string configured in new[] { null, "", name.ToLowerInvariant(), "another" })
                Assert.ThrowsException<InvalidOperationException>(() =>
                    OfficeVbeFixture.RequirePrivateWordDesktop("Word", name, configured, observe));
            Assert.ThrowsException<InvalidOperationException>(() =>
                OfficeVbeFixture.RequirePrivateWordDesktop("Excel", name, name, observe));
            Assert.ThrowsException<InvalidOperationException>(() =>
                OfficeVbeFixture.RequirePrivateWordDesktop("Access", null, name, observe));
            Assert.AreEqual(2, observations, "Refusals must precede desktop and host access.");
            Assert.ThrowsException<InvalidOperationException>(() =>
                OfficeVbeFixture.RequirePrivateWordDesktop("Word", name, name,
                    unused => throw new InvalidOperationException("Inactive desktop not established.")));
        }

        [TestMethod]
        public void NativeWordObjectModelPInvokeMarshalsRequestedIDispatchAsInterface()
        {
            MethodInfo method = typeof(OfficeVbeFixture).GetMethod("NativeAccessibleObjectFromWindow",
                BindingFlags.NonPublic | BindingFlags.Static);
            Assert.IsNotNull(method);
            var parameters = method.GetParameters();
            Assert.AreEqual(4, parameters.Length);
            Assert.AreEqual(typeof(Guid).MakeByRefType(), parameters[2].ParameterType);
            Assert.AreEqual(typeof(object).MakeByRefType(), parameters[3].ParameterType);
            var marshaling = parameters[3].GetCustomAttribute<MarshalAsAttribute>();
            Assert.IsNotNull(marshaling);
            Assert.AreEqual(UnmanagedType.Interface, marshaling.Value);
        }

        [TestMethod]
        public void MacroFreeWordSeedHasOneDocumentPartAndExactRootRelationshipAndRefusesOverwrite()
        {
            string directory = Path.Combine(Path.GetTempPath(), "VBAi-PrivateWordSeed-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(directory);
            string path = Path.Combine(directory, "NativeObjectModelSeed.docx");
            MethodInfo create = typeof(OfficeVbeFixture).GetMethod("WriteMacroFreeSeed", BindingFlags.NonPublic | BindingFlags.Static);
            Assert.IsNotNull(create);
            try
            {
                create.Invoke(null, new object[] { path });
                byte[] original = File.ReadAllBytes(path);
                using (var package = Package.Open(path, FileMode.Open, FileAccess.Read))
                {
                    var parts = package.GetParts().Where(part => !PackUriHelper.IsRelationshipPartUri(part.Uri)).ToArray();
                    Assert.AreEqual(1, parts.Length);
                    Assert.AreEqual("/word/document.xml", parts[0].Uri.ToString());
                    Assert.AreEqual("application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml", parts[0].ContentType);
                    var relations = package.GetRelationships().ToArray();
                    Assert.AreEqual(1, relations.Length);
                    Assert.AreEqual("http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument", relations[0].RelationshipType);
                    Assert.AreEqual(TargetMode.Internal, relations[0].TargetMode);
                    Assert.AreEqual(parts[0].Uri, relations[0].TargetUri);
                    using (var reader = new StreamReader(parts[0].GetStream(), Encoding.UTF8))
                    {
                        string xml = reader.ReadToEnd();
                        StringAssert.Contains(xml, "<w:document");
                        StringAssert.Contains(xml, "<w:body><w:p/></w:body>");
                    }
                }
                var exception = Assert.ThrowsException<TargetInvocationException>(() => create.Invoke(null, new object[] { path }));
                Assert.IsInstanceOfType(exception.InnerException, typeof(AssertFailedException));
                CollectionAssert.AreEqual(original, File.ReadAllBytes(path));
            }
            finally { Directory.Delete(directory, true); }
        }

        [TestMethod]
        public void PrivateWordLaunchRequiresExactReviewedExecutableAndMacroFreeSeedArguments()
        {
            int existsCalls = 0, hashCalls = 0;
            Assert.AreEqual(Path.GetFullPath(Executable), OfficeVbeFixture.RequireExactWordExecutable(
                Executable, Hash, path => { existsCalls++; return path == Executable; },
                path => { hashCalls++; return Hash.ToLowerInvariant(); }));
            Assert.AreEqual(1, existsCalls); Assert.AreEqual(1, hashCalls);
            CollectionAssert.AreEqual(new[] { "/a", @"C:\Owned\NativeObjectModelSeed.docx" },
                OfficeVbeFixture.WordPrivateArguments(@"C:\Owned\NativeObjectModelSeed.docx"));
        }

        [TestMethod]
        public void PrivateWordLaunchRejectsWrongPathHashOrSeedBeforeAnyLauncherCall()
        {
            foreach (string path in new[] { "WINWORD.EXE", @"C:\Owned\Other.exe", @"C:\Owned\WINWORD.EXE\child", "" })
                Assert.ThrowsException<InvalidOperationException>(() => OfficeVbeFixture.RequireExactWordExecutable(
                    path, Hash, unused => { Assert.Fail("No file lookup after an invalid path."); return false; }, unused => Hash));
            foreach (string hash in new[] { null, "", new string('G', 64), new string('A', 63) })
                Assert.ThrowsException<InvalidOperationException>(() => OfficeVbeFixture.RequireExactWordExecutable(
                    Executable, hash, unused => { Assert.Fail("No file lookup after an invalid hash."); return false; }, unused => Hash));
            Assert.ThrowsException<InvalidOperationException>(() => OfficeVbeFixture.RequireExactWordExecutable(
                Executable, Hash, unused => false, unused => { Assert.Fail("No hash after a missing executable."); return Hash; }));
            Assert.ThrowsException<InvalidOperationException>(() => OfficeVbeFixture.RequireExactWordExecutable(
                Executable, Hash, unused => true, unused => new string('B', 64)));
            foreach (string seed in new[] { "seed.docx", @"C:\Owned\Seed.docm", "" })
                Assert.ThrowsException<InvalidOperationException>(() => OfficeVbeFixture.WordPrivateArguments(seed));
        }

        [TestMethod]
        public void PrivateLaunchPersistsOriginalIdentityBeforeDiscoveryAndRequiresActualUiProof()
        {
            string desktop = "VBAiTests_" + Guid.NewGuid().ToString("N");
            foreach (var observation in new[] {
                new IsolatedTestDesktop.ThreadDesktopObservation(41, 0, null, 0, 0),
                new IsolatedTestDesktop.ThreadDesktopObservation(41, 0, null, 5, 0),
                new IsolatedTestDesktop.ThreadDesktopObservation(41, 128, null, 0, 5),
                new IsolatedTestDesktop.ThreadDesktopObservation(41, 128, desktop, 0, 0) })
            {
                var calls = new List<string>();
                OfficeVbeFixture.PreparePrivateWordLaunch(desktop, 41,
                    () => calls.Add("durable-launch"), () => calls.Add("capture-original"), () => calls.Add("inventory"),
                    () => { calls.Add("primary-read"); return observation; },
                    actual => { Assert.AreSame(observation, actual); calls.Add("durable-primary-diagnostic"); },
                    () => {
                        calls.Add("actual-ui-proof");
                        Assert.ThrowsException<InvalidOperationException>(() => VerifyWindow(desktop, read: unused => "Default"));
                        VerifyWindow(desktop);
                        calls.Add("COM-permitted-after-ui-proof");
                    });
                CollectionAssert.AreEqual(new[] { "durable-launch", "capture-original", "inventory", "primary-read",
                    "durable-primary-diagnostic", "actual-ui-proof", "COM-permitted-after-ui-proof" }, calls);
            }
        }

        [TestMethod]
        public void EveryLaunchPhaseFailureStopsWithoutLaterDiscoveryOrComAndPreservesItsOriginalError()
        {
            string desktop = "VBAiTests_" + Guid.NewGuid().ToString("N");
            for (int failAt = 0; failAt < 6; failAt++)
            {
                var calls = new List<int>();
                var original = new InvalidOperationException("Original phase " + failAt);
                Action<int> enter = phase => { calls.Add(phase); if (phase == failAt) throw original; };
                var failure = Assert.ThrowsException<InvalidOperationException>(() =>
                    OfficeVbeFixture.PreparePrivateWordLaunch(desktop, 41,
                        () => enter(0), () => enter(1), () => enter(2),
                        () => { enter(3); return new IsolatedTestDesktop.ThreadDesktopObservation(41, 0, null, 0, 0); },
                        unused => enter(4), () => enter(5)));
                Assert.AreSame(original, failure);
                CollectionAssert.AreEqual(Enumerable.Range(0, failAt + 1).ToArray(), calls);
            }
        }

        [TestMethod]
        public void InvalidLaunchDependenciesOrGenerationNeverReachCom()
        {
            string desktop = "VBAiTests_" + Guid.NewGuid().ToString("N");
            for (int missing = 0; missing < 6; missing++)
            {
                Action noAccess = () => Assert.Fail("Invalid dependencies cannot access launch evidence or COM.");
                Func<IsolatedTestDesktop.ThreadDesktopObservation> read = () => { noAccess(); return null; };
                Assert.ThrowsException<ArgumentNullException>(() => OfficeVbeFixture.PreparePrivateWordLaunch(desktop, 41,
                    missing == 0 ? null : noAccess, missing == 1 ? null : noAccess, missing == 2 ? null : noAccess,
                    missing == 3 ? null : read, missing == 4 ? null : new Action<IsolatedTestDesktop.ThreadDesktopObservation>(unused => noAccess()),
                    missing == 5 ? null : noAccess));
            }
            foreach (var observation in new[] {
                null,
                new IsolatedTestDesktop.ThreadDesktopObservation(42, 128, desktop, 0, 0),
                new IsolatedTestDesktop.ThreadDesktopObservation(41, 128, "Default", 0, 0),
                new IsolatedTestDesktop.ThreadDesktopObservation(41, 0, desktop, 0, 0) })
            {
                int durableLaunches = 0, captures = 0, attaches = 0;
                Assert.ThrowsException<InvalidOperationException>(() => OfficeVbeFixture.PreparePrivateWordLaunch(desktop, 41,
                    () => durableLaunches++, () => captures++, () => { }, () => observation, unused => { }, () => attaches++));
                Assert.AreEqual(1, durableLaunches); Assert.AreEqual(1, captures); Assert.AreEqual(0, attaches);
            }
            Assert.ThrowsException<ArgumentException>(() => OfficeVbeFixture.PreparePrivateWordLaunch("Default", 41,
                () => Assert.Fail(), () => Assert.Fail(), () => Assert.Fail(), () => null, unused => Assert.Fail(), () => Assert.Fail()));
            Assert.ThrowsException<ArgumentException>(() => OfficeVbeFixture.PreparePrivateWordLaunch(desktop, 0,
                () => Assert.Fail(), () => Assert.Fail(), () => Assert.Fail(), () => null, unused => Assert.Fail(), () => Assert.Fail()));
        }

        [TestMethod]
        public void CompleteWordWindowIdentityMustPrecedeTheActualUiDesktopRead()
        {
            string desktop = "VBAiTests_" + Guid.NewGuid().ToString("N");
            int desktopReads = 0;
            Func<uint, string> read = tid => { Assert.AreEqual(41u, tid); desktopReads++; return desktop; };
            VerifyWindow(desktop, read: read);
            Assert.AreEqual(1, desktopReads);
            foreach (Action reject in new Action[] {
                () => VerifyWindow(desktop, handle: IntPtr.Zero, read: read),
                () => VerifyWindow(desktop, root: IntPtr.Zero, read: read),
                () => VerifyWindow(desktop, root: new IntPtr(10), read: read),
                () => VerifyWindow(desktop, pid: 8, read: read),
                () => VerifyWindow(desktop, rootPid: 8, read: read),
                () => VerifyWindow(desktop, expectedPid: 0, read: read),
                () => VerifyWindow(desktop, tid: 0, read: read),
                () => VerifyWindow(desktop, rootTid: 42, read: read),
                () => VerifyWindow(desktop, childClass: "Other", read: read),
                () => VerifyWindow(desktop, rootClass: "Other", read: read),
                () => VerifyWindow(desktop, visible: false, read: read),
                () => VerifyWindow(desktop, rootVisible: false, read: read) })
                Assert.ThrowsException<InvalidOperationException>(reject);
            Assert.AreEqual(1, desktopReads, "Foreign/stale window metadata must refuse before any thread desktop read.");
            foreach (string actual in new[] { null, "", "Default", desktop.ToUpperInvariant() })
                Assert.ThrowsException<InvalidOperationException>(() => VerifyWindow(desktop, read: unused => actual));
            var original = new InvalidOperationException("Actual UI desktop unavailable");
            Assert.AreSame(original, Assert.ThrowsException<InvalidOperationException>(() =>
                VerifyWindow(desktop, read: unused => throw original)));
            Assert.ThrowsException<ArgumentException>(() => VerifyWindow("Default", read: read));
            Assert.ThrowsException<ArgumentNullException>(() => OfficeVbeFixture.RequirePrivateWordWindowIdentity(
                new IntPtr(10), new IntPtr(20), 7, 7, 41, 41, "_WwG", "OpusApp", true, true, 7, desktop, null));
        }

        private static void VerifyWindow(string desktop, IntPtr? handle = null, IntPtr? root = null, uint pid = 7,
            uint rootPid = 7, uint tid = 41, uint rootTid = 41, string childClass = "_WwG", string rootClass = "OpusApp",
            bool visible = true, bool rootVisible = true, uint expectedPid = 7, Func<uint, string> read = null)
        {
            OfficeVbeFixture.RequirePrivateWordWindowIdentity(handle ?? new IntPtr(10), root ?? new IntPtr(20),
                pid, rootPid, tid, rootTid, childClass, rootClass, visible, rootVisible, expectedPid, desktop, read ?? (unused => desktop));
        }

        [TestMethod]
        public void LateUniqueWindowIsReadOnlyObservedAndVerifiedBeforeReturning()
        {
            int inventories = 0, waits = 0, checks = 0, verifies = 0;
            IntPtr handle = new IntPtr(10);
            Assert.AreEqual(handle, OfficeVbeFixture.WaitForPrivateWordWindow(
                () => { checks++; return false; },
                () => ++inventories == 1 ? new IntPtr[0] : new[] { handle }, () => true, () => waits++,
                actual => { Assert.AreEqual(handle, actual); verifies++; }));
            Assert.AreEqual(2, inventories); Assert.AreEqual(1, waits);
            Assert.AreEqual(3, checks); Assert.AreEqual(1, verifies);
        }

        [TestMethod]
        public void AmbiguousIncompleteOrFailedWindowInventoryCannotReachVerificationOrRetry()
        {
            foreach (IntPtr[] windows in new[] { null, new[] { IntPtr.Zero },
                new[] { new IntPtr(10), new IntPtr(20) }, new[] { new IntPtr(10), new IntPtr(10) } })
            {
                int inventories = 0;
                Assert.ThrowsException<InvalidOperationException>(() => OfficeVbeFixture.WaitForPrivateWordWindow(() => false,
                    () => { inventories++; return windows; }, () => true, () => Assert.Fail("No inventory replay."),
                    unused => Assert.Fail("No UI verification/COM after ambiguous inventory.")));
                Assert.AreEqual(1, inventories);
            }
            var original = new InvalidOperationException("Named desktop enumeration failed or exceeded its bound");
            Assert.AreSame(original, Assert.ThrowsException<InvalidOperationException>(() => OfficeVbeFixture.WaitForPrivateWordWindow(
                () => false, () => throw original, () => true, () => Assert.Fail(), unused => Assert.Fail())));
        }

        [TestMethod]
        public void OriginalExitUiVerificationFailureAndDeadlineStopWithoutComOrFurtherObservation()
        {
            Assert.ThrowsException<InvalidOperationException>(() => OfficeVbeFixture.WaitForPrivateWordWindow(() => true,
                () => { Assert.Fail(); return null; }, () => true, () => Assert.Fail(), unused => Assert.Fail()));
            int checks = 0, verifies = 0;
            Assert.ThrowsException<InvalidOperationException>(() => OfficeVbeFixture.WaitForPrivateWordWindow(() => ++checks == 2,
                () => new[] { new IntPtr(10) }, () => true, () => Assert.Fail(), unused => verifies++));
            Assert.AreEqual(2, checks); Assert.AreEqual(1, verifies);
            var original = new InvalidOperationException("Stale/foreign/UI desktop refusal");
            Assert.AreSame(original, Assert.ThrowsException<InvalidOperationException>(() => OfficeVbeFixture.WaitForPrivateWordWindow(
                () => false, () => new[] { new IntPtr(10) }, () => true, () => Assert.Fail(), unused => throw original)));
            Assert.ThrowsException<TimeoutException>(() => OfficeVbeFixture.WaitForPrivateWordWindow(
                () => { Assert.Fail(); return false; }, () => { Assert.Fail(); return null; }, () => false, () => Assert.Fail(), unused => Assert.Fail()));
            int bounds = 0, waits = 0;
            Assert.ThrowsException<TimeoutException>(() => OfficeVbeFixture.WaitForPrivateWordWindow(() => false,
                () => new IntPtr[0], () => ++bounds == 1, () => waits++, unused => Assert.Fail()));
            Assert.AreEqual(1, waits); Assert.AreEqual(2, bounds);
            Assert.ThrowsException<ArgumentNullException>(() => OfficeVbeFixture.WaitForPrivateWordWindow(null, null, null, null, null));
        }

        [TestMethod]
        public void WordDocumentInventoryCollectsZeroOneTwoAndDeduplicatedOwnedChildrenAndIgnoresForeignPids()
        {
            foreach (int count in new[] { 0, 1, 2 })
            {
                var expected = Enumerable.Range(10, count).Select(value => new IntPtr(value)).ToArray();
                int childrenReads = 0, classes = 0;
                var actual = OfficeVbeFixture.CollectPrivateWordDocumentWindows(7,
                    visitor => { Assert.IsTrue(visitor(new IntPtr(1))); Assert.IsTrue(visitor(new IntPtr(2))); },
                    (root, visitor) => {
                        Assert.AreEqual(new IntPtr(1), root); childrenReads++;
                        foreach (var child in expected) { Assert.IsTrue(visitor(child)); Assert.IsTrue(visitor(child)); }
                        Assert.IsTrue(visitor(new IntPtr(30))); Assert.IsTrue(visitor(new IntPtr(31)));
                    }, handle => handle.ToInt64() == 2 || handle.ToInt64() == 30 ? 8u : 7u,
                    handle => { Assert.AreNotEqual(new IntPtr(30), handle); classes++; return handle.ToInt64() == 31 ? "Other" : "_WwG"; });
                CollectionAssert.AreEquivalent(expected, actual);
                Assert.AreEqual(1, childrenReads); Assert.AreEqual(count * 2 + 1, classes);
            }
        }

        [TestMethod]
        public void WordDocumentInventoryRefusesEveryBoundAndKeepsDiscoveryFailuresBeforeAnyCom()
        {
            foreach (int bound in new[] { 0, 1, 2 })
            {
                Assert.ThrowsException<InvalidOperationException>(() => OfficeVbeFixture.CollectPrivateWordDocumentWindows(7,
                    visitor => {
                        int roots = bound == 0 ? 4097 : 1;
                        for (int index = 1; index <= roots; index++) if (!visitor(new IntPtr(index))) break;
                    },
                    (root, visitor) => {
                        if (bound == 0) return;
                        for (int index = 1; index <= (bound == 1 ? 2049 : 9); index++) if (!visitor(new IntPtr(10000 + index))) break;
                    }, unused => 7, unused => bound == 1 ? "Other" : "_WwG"));
            }
            var original = new InvalidOperationException("Original discovery failure");
            for (int failure = 0; failure < 4; failure++)
            {
                Assert.AreSame(original, Assert.ThrowsException<InvalidOperationException>(() =>
                    OfficeVbeFixture.CollectPrivateWordDocumentWindows(7,
                        visitor => { if (failure == 0) throw original; visitor(new IntPtr(1)); },
                        (root, visitor) => { if (failure == 1) throw original; visitor(new IntPtr(10)); },
                        unused => { if (failure == 2) throw original; return 7; },
                        unused => { if (failure == 3) throw original; return "_WwG"; })));
            }
            Assert.ThrowsException<ArgumentNullException>(() => OfficeVbeFixture.CollectPrivateWordDocumentWindows(7, null, null, null, null));
            Assert.ThrowsException<ArgumentException>(() => OfficeVbeFixture.CollectPrivateWordDocumentWindows(0, null, null, null, null));
        }
    }
}
