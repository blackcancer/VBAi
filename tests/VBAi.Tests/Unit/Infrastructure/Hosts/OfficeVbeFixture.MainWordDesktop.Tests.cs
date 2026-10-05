using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class OfficeVbeFixtureMainWordDesktopTests
    {
        [DataTestMethod]
        [DataRow(0)] [DataRow(1)] [DataRow(2)] [DataRow(3)] [DataRow(4)] [DataRow(5)] [DataRow(6)]
        [DataRow(7)] [DataRow(8)] [DataRow(9)] [DataRow(10)] [DataRow(11)]
        public void ExactObservedRootIsRequiredAndHiddenStartupIsOnlyObservation(int fault)
        {
            var observed = IsolatedTestDesktopMainTests.Inventory().Windows[0];
            IntPtr root = new IntPtr(11); uint pid = 42, tid = 7; string expectedClass = "OpusApp"; bool requireVisible = true;
            switch (fault)
            {
                case 1: observed = null; break; case 2: root = IntPtr.Zero; break;
                case 3: pid = 0; break; case 4: tid = 0; break; case 5: observed.Handle = 13; break;
                case 6: observed.ProcessId = 43; break; case 7: observed.ThreadId = 8; break;
                case 8: expectedClass = null; break; case 9: observed.ClassName = "Other"; break;
                case 10: observed.Visible = false; break;
                case 11: observed.Visible = false; requireVisible = false; break;
            }
            Action run = () => OfficeVbeFixture.RequireMainWordWindowObservation(observed, root, pid, tid, requireVisible, expectedClass);
            if (fault == 0 || fault == 11) run(); else Assert.ThrowsException<InvalidOperationException>(run);
        }

        [DataTestMethod]
        [DataRow(null)] [DataRow("")] [DataRow("0")] [DataRow("true")] [DataRow("1 ")] [DataRow(" 1")]
        public void MainSelectionRequiresTheExactOptInAndDoesNotInvokeItsGuardOtherwise(string optIn)
        {
            int guard = 0;
            Assert.IsFalse(OfficeVbeFixture.SelectMainWordDesktop("Word", optIn, null, null, () => guard++));
            Assert.IsFalse(OfficeVbeFixture.SelectMainWordDesktop("Word", optIn, "Private", "Private", () => guard++));
            Assert.AreEqual(0, guard);
        }

        [DataTestMethod]
        [DataRow("Word", null, null, true)] [DataRow("Word", "", "", true)]
        [DataRow("Word", "Private", "Private", false)] [DataRow("Word", null, "Default", false)]
        [DataRow("Word", "Default", null, false)] [DataRow("Word", " ", null, false)]
        [DataRow("PowerPoint", null, null, false)] [DataRow(null, null, null, false)]
        public void MainSelectionNeverOverridesPrivateSelectionOrAnotherHost(string kind, string required, string configured, bool valid)
        {
            int count = 0;
            Func<bool> call = () => OfficeVbeFixture.SelectMainWordDesktop(kind, "1", required, configured, () => count++);
            if (valid) { Assert.IsTrue(call()); Assert.AreEqual(1, count); }
            else { Assert.ThrowsException<InvalidOperationException>(() => call()); Assert.AreEqual(0, count); }
        }

        [TestMethod]
        public void MainSelectionPreservesPlacementFailureAndPrivateSelectionStillUsesItsOwnGuard()
        {
            var error = new InvalidOperationException("input changed");
            Assert.AreSame(error, Assert.ThrowsException<InvalidOperationException>(() =>
                OfficeVbeFixture.SelectMainWordDesktop("Word", "1", null, null, () => { throw error; })));
            int privateChecks = 0;
            Assert.AreEqual("Private", OfficeVbeFixture.RequirePrivateWordDesktop("Word", "Private", "Private", name => privateChecks++));
            Assert.AreEqual(1, privateChecks);
        }

        [DataTestMethod]
        [DataRow(-1)] [DataRow(0)] [DataRow(1)] [DataRow(2)] [DataRow(3)] [DataRow(4)] [DataRow(5)]
        public void MainLaunchPersistsTheOriginalBeforeAnyAttachAndStopsAtEveryFailure(int fault)
        {
            var log = new List<string>(); int step = 0; var error = new InvalidOperationException("preserved");
            Action<string> action = name => { log.Add(name); if (step++ == fault) throw error; };
            Action run = () => OfficeVbeFixture.PrepareMainWordLaunch(() => action("receipt"), () => action("capture"),
                () => action("inventory"), () => action("placement"), () => action("attach"));
            if (fault < 0) run(); else Assert.AreSame(error, Assert.ThrowsException<InvalidOperationException>(run));
            string[] expected = { "receipt", "capture", "placement", "inventory", "placement", "attach" };
            CollectionAssert.AreEqual(expected.Take(fault < 0 ? 6 : fault + 1).ToArray(), log);
        }

        [DataTestMethod]
        [DataRow(-1)] [DataRow(0)] [DataRow(1)] [DataRow(2)] [DataRow(3)] [DataRow(4)] [DataRow(5)]
        [DataRow(6)] [DataRow(7)]
        public void SeedTransitionPreservesUncertainCreateCloseAndReleaseWithoutReplay(int fault)
        {
            var log = new List<string>(); int guards = 0; bool newDocument = false, closed = false, released = false, cleared = false;
            var error = new InvalidOperationException("uncertain original");
            Action run = () => OfficeVbeFixture.ReplaceWordSeedOnce(
                () => { log.Add("guard"); if (fault == (guards++ == 0 ? 0 : 3)) throw error; },
                () => { log.Add("create"); if (fault == 1) throw error; newDocument = true; if (fault == 2) throw error; },
                () => { log.Add("close"); if (fault == 4) throw error; closed = true; if (fault == 5) throw error; },
                () => { log.Add("release"); released = true; if (fault == 6) throw error; },
                () => { log.Add("clear"); if (fault == 7) throw error; cleared = true; });
            if (fault < 0) run(); else Assert.AreSame(error, Assert.ThrowsException<InvalidOperationException>(run));
            Assert.AreEqual(fault < 0 || fault >= 2, newDocument);
            Assert.AreEqual(fault < 0 || fault >= 5, closed);
            Assert.AreEqual(fault < 0 || fault >= 6, released);
            Assert.AreEqual(fault < 0, cleared);
            Assert.IsTrue(log.Count(x => x == "create") <= 1 && log.Count(x => x == "close") <= 1 && log.Count(x => x == "release") <= 1);
        }

        [DataTestMethod]
        [DataRow(0)] [DataRow(1)] [DataRow(2)] [DataRow(3)] [DataRow(4)] [DataRow(5)]
        public void PlacementObservationChecksOriginalBeforeAndAfterAndNeverRetries(int fault)
        {
            int guards = 0, reads = 0; var original = new InvalidOperationException("original generation changed");
            Func<IsolatedTestDesktop.MainInventory> run = () => OfficeVbeFixture.ObserveMainWordPlacement(
                () => { guards++; if ((fault == 1 && guards == 1) || (fault == 2 && guards == 2)) throw original; },
                () => { reads++; if (fault == 3) throw original; var data = IsolatedTestDesktopMainTests.Inventory();
                    if (fault == 4) data.Complete = false; if (fault == 5) data.Windows[0].ProcessId = 99; return data; },
                42, new IntPtr(11), true, true);
            if (fault == 0) Assert.IsNotNull(run());
            else { var caught = Assert.ThrowsException<InvalidOperationException>(() => run()); if (fault <= 3) Assert.AreSame(original, caught); }
            Assert.AreEqual(fault == 1 ? 0 : 1, reads);
            Assert.AreEqual(fault == 0 || fault == 2 ? 2 : 1, guards);
        }

        [DataTestMethod]
        [DataRow(0)] [DataRow(1)] [DataRow(2)] [DataRow(3)] [DataRow(4)] [DataRow(5)] [DataRow(6)] [DataRow(7)] [DataRow(8)] [DataRow(9)] [DataRow(10)]
        public void OriginalNativeHandleAndCapturedGenerationMustBothRemainExact(int fault)
        {
            Action run = () => OfficeVbeFixture.RequireMainWordGeneration(fault == 1 ? 0 : 42,
                fault == 2 ? null : "2026-10-05T19:00:00.0000000Z", fault == 3 ? null : @"C:\Office\WINWORD.EXE",
                fault == 4 ? IntPtr.Zero : new IntPtr(51), fault == 5 ? 43 : 42, fault == 6, fault == 10 ? 43 : 42, fault == 7,
                fault == 8 ? "2026-10-05T19:00:01.0000000Z" : "2026-10-05T19:00:00.0000000Z",
                fault == 9 ? @"C:\Foreign\WINWORD.EXE" : @"c:\office\winword.exe");
            if (fault == 0) run(); else Assert.ThrowsException<InvalidOperationException>(run);
        }

        [DataTestMethod]
        [DataRow(0)] [DataRow(1)] [DataRow(2)] [DataRow(3)] [DataRow(4)] [DataRow(5)] [DataRow(6)]
        [DataRow(7)] [DataRow(8)] [DataRow(9)] [DataRow(10)] [DataRow(11)] [DataRow(12)] [DataRow(13)]
        public void ReadyProductUsesActualStatusWithFrozenMvidAndFileHash(int fault)
        {
            Guid mvid = Guid.NewGuid(); string pin = mvid.ToString("D"), sha = new string('A', 64);
            var status = new Dictionary<string, object> { ["HostProcessId"] = 42, ["AssemblyPath"] = @"C:\Candidate\VBAi.dll", ["AssemblyModuleVersionId"] = pin };
            switch (fault)
            {
                case 1: status = null; break; case 2: status.Remove("HostProcessId"); break;
                case 3: status["HostProcessId"] = "42"; break; case 4: status["HostProcessId"] = 43; break;
                case 5: status.Remove("AssemblyPath"); break; case 6: status["AssemblyModuleVersionId"] = Guid.NewGuid().ToString("D"); break;
                case 7: pin = null; break; case 8: pin = Guid.NewGuid().ToString("D"); break;
                case 9: sha = new string('X', 64); break; case 10: sha = "A"; break;
                case 11: status["AssemblyPath"] = "relative.dll"; break;
            }
            int reads = 0; var readError = new InvalidOperationException("file read failed");
            Func<string> run = () => OfficeVbeFixture.RequireMainWordProduct(status, 42, pin, sha, mvid, path => {
                reads++; Assert.AreEqual(@"C:\Candidate\VBAi.dll", path); if (fault == 13) throw readError;
                return fault == 12 ? new string('B', 64) : sha.ToLowerInvariant(); });
            if (fault == 0) Assert.AreEqual(@"C:\Candidate\VBAi.dll", run());
            else if (fault == 11) Assert.ThrowsException<ArgumentException>(() => run());
            else { var error = Assert.ThrowsException<InvalidOperationException>(() => run()); if (fault == 13) Assert.AreSame(readError, error); }
            Assert.AreEqual(fault == 0 || fault >= 12 ? 1 : 0, reads);
        }
    }
}
