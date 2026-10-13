using Microsoft.CSharp;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.CodeDom.Compiler;
using System.IO;
using System.Linq;
using System.Windows.Automation;

namespace VBAi.Tests.Unit
{
    /// <summary>Complete admission, independent identity readback, one-shot delivery and uncertain-completion matrix without an Office/CAD host.</summary>
    [TestClass, TestCategory("Unit")]
    public sealed class SolidWorksTaskDialogTeardownTests
    {
        private static Q020TaskDialogTeardown.Scope Scope() => new Q020TaskDialogTeardown.Scope(42, 73, 100, 200, @"E:\Qualification\Owned.swp");
        private static Q020TaskDialogTeardown.Snapshot Snapshot(long dialog = 1000, bool french = true, bool buttonType = false)
        {
            string[] labels = french ? new[] { "Oui", "Non", "Annuler" } : new[] { "Yes", "No", "Cancel" };
            return new Q020TaskDialogTeardown.Snapshot
            {
                Dialog = dialog, OwnerRoot = 200, Pid = 42, Thread = 73, Class = "#32770", Caption = "SOLIDWORKS", Visible = true, Enabled = true,
                Elements = new[]
                {
                    new Q020TaskDialogTeardown.Element { Id = "MainIcon", Name = "MainInstructionIcon", Type = "ControlType.Image", Pid = 42, Enabled = true },
                    new Q020TaskDialogTeardown.Element { Id = "ContentText", Name = french ? "Enregistrer les modifications apportées à " + Scope().Macro + " ?" : "Save changes to " + Scope().Macro + "?", Type = "ControlType.Text", Pid = 42, Enabled = true },
                    Command("6", labels[0], dialog + 1, buttonType), Command("7", labels[1], dialog + 2, buttonType), Command("2", labels[2], dialog + 3, buttonType),
                    new Q020TaskDialogTeardown.Element { Type = "ControlType.Pane", Pid = 42, Enabled = true }
                },
                Buttons = Enumerable.Range(0, 3).Select(i => new Q020TaskDialogTeardown.Button
                { Handle = dialog + i + 1, Root = dialog, Pid = 42, Thread = 73, Id = 0, Class = "Button", Caption = i == 2 ? labels[i] : "&" + labels[i], Visible = true, Enabled = true }).ToArray()
            };
        }
        private static Q020TaskDialogTeardown.Element Command(string id, string label, long handle, bool button) => new Q020TaskDialogTeardown.Element
        { Id = "CommandButton_" + id, Name = label, Handle = handle, Type = button ? "ControlType.Button" : "ControlType.Pane", Pid = 42, Enabled = true, Actionable = true };

        private static Q020TaskDialogTeardown.Snapshot WithChrome(bool french = true)
        {
            var snapshot = Snapshot(1000, french, true);
            snapshot.Elements[0].Actionable = true;
            var ids = new[] { "TitleBar", "SystemMenuBar", "Item 1", "Close" };
            var names = french ? new[] { "SOLIDWORKS", "Barre de menus système", "Système", "Fermer" }
                : new[] { "SOLIDWORKS", "System menu bar", "System", "Close" };
            var types = new[] { "TitleBar", "MenuBar", "MenuItem", "Button" };
            snapshot.Elements = snapshot.Elements.Concat(Enumerable.Range(0, 4).Select(i => new Q020TaskDialogTeardown.Element
            { Id = ids[i], Name = names[i], Type = "ControlType." + types[i], Pid = 42, Enabled = true, Offscreen = i == 1, Actionable = i >= 2 })).ToArray();
            return snapshot;
        }

        [DataTestMethod, DataRow(false), DataRow(true)]
        public void ObservedNonclientChromeAndInstructionValuePatternNeverBecomeMacroCommands(bool french)
        {
            var original = WithChrome(french); long target = 0;
            var batch = new Q020TaskDialogTeardown.DiscardBatch(Scope(), new[] { original }, () => new[] { WithChrome(french) }, () => { }, h => { target = h; return true; });
            batch.RequestNext(); Assert.AreEqual(1002L, target);
            Assert.IsFalse(batch.ObserveCompletion()); Assert.AreEqual(1000L, batch.PendingDialog);
        }

        [DataTestMethod]
        [DataRow("incomplete")][DataRow("duplicate")][DataRow("extra")][DataRow("name")][DataRow("type")]
        [DataRow("handle")][DataRow("disabled")][DataRow("visibility")][DataRow("action")][DataRow("pid")][DataRow("language")]
        [DataRow("icon duplicate")][DataRow("icon handle")][DataRow("icon disabled")][DataRow("icon hidden")]
        public void EveryChangedChromeOrInstructionBindingRefusesBeforeAnyNativeDelivery(string fault)
        {
            var snapshot = WithChrome(); int posts = 0;
            var node = snapshot.Elements.Last();
            if (fault == "incomplete") snapshot.Elements = snapshot.Elements.Take(snapshot.Elements.Length - 1).ToArray();
            if (fault == "duplicate") node.Id = "TitleBar";
            if (fault == "extra") snapshot.Elements = snapshot.Elements.Concat(new[] { node }).ToArray();
            if (fault == "name") node.Name = "Unexpected"; if (fault == "type") node.Type = "ControlType.Pane";
            if (fault == "handle") node.Handle = 9000; if (fault == "disabled") node.Enabled = false;
            if (fault == "visibility") node.Offscreen = true; if (fault == "action") node.Actionable = false;
            if (fault == "pid") node.Pid = 43; if (fault == "language") node.Name = "Close";
            if (fault == "icon duplicate") snapshot.Elements = snapshot.Elements.Concat(new[] { snapshot.Elements[0] }).ToArray();
            if (fault == "icon handle") snapshot.Elements[0].Handle = 9000;
            if (fault == "icon disabled") snapshot.Elements[0].Enabled = false;
            if (fault == "icon hidden") snapshot.Elements[0].Offscreen = true;
            var batch = new Q020TaskDialogTeardown.DiscardBatch(Scope(), new[] { WithChrome() }, () => new[] { snapshot }, () => { }, _ => { posts++; return true; });
            Assert.ThrowsException<InvalidOperationException>(batch.RequestNext);
            Assert.ThrowsException<InvalidOperationException>(batch.RequestNext); Assert.AreEqual(0, posts);
        }

        [DataTestMethod, DataRow(false), DataRow(true)]
        public void OriginalNormalOwnerExitCanSettleOnlyTheLastCapturedQueuedQuestion(bool two)
        {
            var current = two ? new[] { Snapshot(), Snapshot(2000) } : new[] { Snapshot() }; int posts = 0, proofs = 0;
            var batch = new Q020TaskDialogTeardown.DiscardBatch(Scope(), current, () => current, () => { }, _ => { posts++; return true; });
            Assert.ThrowsException<ArgumentNullException>(() => batch.ObserveOriginalOwnerExit(null));
            Assert.ThrowsException<InvalidOperationException>(() => batch.ObserveOriginalOwnerExit(() => true));
            batch.RequestNext();
            if (two)
            {
                Assert.ThrowsException<InvalidOperationException>(() => batch.ObserveOriginalOwnerExit(() => { proofs++; return true; }));
                Assert.AreEqual(0, proofs); current = new[] { Snapshot(2000) }; Assert.IsTrue(batch.ObserveCompletion()); batch.RequestNext();
            }
            Assert.IsFalse(batch.ObserveOriginalOwnerExit(() => false));
            Assert.ThrowsException<InvalidOperationException>(() => batch.ObserveOriginalOwnerExit(() => { throw new InvalidOperationException("Exit receipt unavailable"); }));
            Assert.ThrowsException<InvalidOperationException>(batch.RequestNext);
            Assert.IsTrue(batch.ObserveOriginalOwnerExit(() => true)); Assert.IsTrue(batch.Complete);
            Assert.AreEqual(two ? 2 : 1, posts);
            Assert.ThrowsException<InvalidOperationException>(() => batch.ObserveOriginalOwnerExit(() => true));
        }

        [DataTestMethod, DataRow(false), DataRow(true)]
        public void OriginalOwnerExitCannotPromoteAnUncertainOrFailedDelivery(bool throws)
        {
            int proofs = 0;
            var batch = new Q020TaskDialogTeardown.DiscardBatch(Scope(), new[] { Snapshot() }, () => new[] { Snapshot() }, () => { }, _ =>
            { if (throws) throw new InvalidOperationException("Unknown delivery"); return false; });
            Assert.ThrowsException<InvalidOperationException>(batch.RequestNext);
            Assert.ThrowsException<InvalidOperationException>(() => batch.ObserveOriginalOwnerExit(() => { proofs++; return true; }));
            Assert.AreEqual(0, proofs); Assert.IsFalse(batch.Complete);
            Assert.IsFalse(batch.LastOriginalQuestionPending);
        }

        [DataTestMethod, DataRow(false), DataRow(true)]
        public void QuietAdmissionRequiresElapsedTimeAndThreeReadsForQuestionsOrAnEmptyBoundary(bool empty)
        {
            var observer = new Q020TaskDialogTeardown.StableInventory(Scope(), 1000);
            var snapshots = empty ? new Q020TaskDialogTeardown.Snapshot[0] : new[] { WithChrome() };
            Assert.IsFalse(observer.Observe(snapshots, 0)); Assert.IsFalse(observer.Observe(snapshots, 1000));
            Assert.IsTrue(observer.Observe(snapshots, 1001)); Assert.AreEqual(!empty, observer.HasQuestions);
        }

        [DataTestMethod]
        [DataRow("added")][DataRow("removed")][DataRow("replaced")][DataRow("reordered")][DataRow("command")][DataRow("empty to question")]
        public void EveryValidInventoryChangeRestartsTheQuietBoundaryBeforeAnyInput(string change)
        {
            var observer = new Q020TaskDialogTeardown.StableInventory(Scope(), 1000);
            var snapshots = change == "empty to question" ? new Q020TaskDialogTeardown.Snapshot[0]
                : change == "added" ? new[] { Snapshot() } : new[] { Snapshot(), Snapshot(2000) };
            observer.Observe(snapshots, 0); observer.Observe(snapshots, 500); Assert.IsTrue(observer.Observe(snapshots, 1000));
            if (change == "added") snapshots = new[] { Snapshot(), Snapshot(3000) };
            if (change == "removed") snapshots = new[] { Snapshot() };
            if (change == "replaced") snapshots = new[] { Snapshot(3000), Snapshot(4000) };
            if (change == "reordered") snapshots = snapshots.Reverse().ToArray();
            if (change == "command") { snapshots[0].Buttons[1].Handle = 9000; snapshots[0].Elements[3].Handle = 9000; }
            if (change == "empty to question") snapshots = new[] { Snapshot() };
            Assert.IsFalse(observer.Observe(snapshots, 1100)); Assert.IsFalse(observer.Observe(snapshots, 1500));
            Assert.IsFalse(observer.Observe(snapshots, 2099)); Assert.IsTrue(observer.Observe(snapshots, 2100));
        }

        [DataTestMethod, DataRow("null"), DataRow("foreign"), DataRow("clock"), DataRow("negative")]
        public void InvalidReadNeverLeavesAStaleReadyInventory(string fault)
        {
            var observer = new Q020TaskDialogTeardown.StableInventory(Scope(), 1000);
            observer.Observe(new[] { Snapshot() }, 0); observer.Observe(new[] { Snapshot() }, 500); observer.Observe(new[] { Snapshot() }, 1000);
            var snapshots = fault == "null" ? null : new[] { Snapshot() };
            if (fault == "foreign") snapshots[0].Pid = 43;
            long time = fault == "clock" ? 999 : fault == "negative" ? -1 : 1500;
            Assert.ThrowsException<InvalidOperationException>(() => observer.Observe(snapshots, time)); Assert.IsFalse(observer.Ready);
            Assert.IsFalse(observer.Observe(new[] { Snapshot() }, 2000));
            Assert.IsFalse(observer.Observe(new[] { Snapshot() }, 2999));
            Assert.IsTrue(observer.Observe(new[] { Snapshot() }, 3000));
        }

        [TestMethod]
        public void AdmissionConfigurationAndLastQuestionStatusPreserveTheirStrictBoundaries()
        {
            Assert.ThrowsException<ArgumentNullException>(() => new Q020TaskDialogTeardown.StableInventory(null, 1000));
            Assert.ThrowsException<ArgumentOutOfRangeException>(() => new Q020TaskDialogTeardown.StableInventory(Scope(), 499));
            Assert.ThrowsException<ArgumentOutOfRangeException>(() => new Q020TaskDialogTeardown.StableInventory(Scope(), 5001));
            var observer = new Q020TaskDialogTeardown.StableInventory(Scope(), 5000);
            Assert.IsFalse(observer.Ready); Assert.IsFalse(observer.HasQuestions);
            var current = new[] { Snapshot(), Snapshot(2000) };
            var batch = new Q020TaskDialogTeardown.DiscardBatch(Scope(), current, () => current, () => { }, _ => true);
            Assert.IsFalse(batch.LastOriginalQuestionPending); batch.RequestNext(); Assert.IsFalse(batch.LastOriginalQuestionPending);
            current = new[] { Snapshot(2000) }; batch.ObserveCompletion(); batch.RequestNext(); Assert.IsTrue(batch.LastOriginalQuestionPending);
            batch.ObserveOriginalOwnerExit(() => true); Assert.IsFalse(batch.LastOriginalQuestionPending);
        }

        [TestMethod]
        public void UnavailableObservationCannotAdvanceOrRedeliverAndBreaksTheQuietInterval()
        {
            bool unavailable = false; int posts = 0;
            var current = new[] { Snapshot(), Snapshot(2000) };
            var batch = new Q020TaskDialogTeardown.DiscardBatch(Scope(), current, () =>
            { if (unavailable) throw new Q020TaskDialogTeardown.ObservationUnavailableException("Control transitioning"); return current; }, () => { }, _ => { posts++; return true; });
            batch.RequestNext(); unavailable = true;
            Assert.ThrowsException<Q020TaskDialogTeardown.ObservationUnavailableException>(() => batch.ObserveCompletion());
            Assert.ThrowsException<InvalidOperationException>(batch.RequestNext); Assert.AreEqual(1, posts); Assert.AreEqual(0, batch.Settled);
            unavailable = false; current = new[] { Snapshot(2000) }; Assert.IsTrue(batch.ObserveCompletion());
            batch.RequestNext(); Assert.AreEqual(2, posts);
            var observer = new Q020TaskDialogTeardown.StableInventory(Scope(), 1000);
            observer.Observe(current, 0); observer.Observe(current, 500); Assert.IsTrue(observer.Observe(current, 1000));
            observer.Invalidate(); Assert.IsFalse(observer.Ready); Assert.IsFalse(observer.HasQuestions);
            Assert.IsFalse(observer.Observe(current, 2000)); Assert.IsFalse(observer.Observe(current, 2999)); Assert.IsTrue(observer.Observe(current, 3000));
        }

        private const string ReaderNonce = "0123456789abcdef0123456789abcdef";
        private const string ReaderHash = "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA";
        private const string ReaderBirth = "2026-10-09T13:00:00.0000000Z";
        private static string ReaderManifest => @"E:\Qualification\" + ReaderNonce + ".request.json";
        private static Q020TaskDialogTeardown.ReaderRequest ReaderRequest() => new Q020TaskDialogTeardown.ReaderRequest
        {
            Version = 1, Nonce = ReaderNonce, UserSid = "S-1-5-21-42", ParentPid = 41, HostPid = 42, Thread = 73, VbeRoot = 100, NativeRoot = 200,
            ParentBirthUtc = ReaderBirth, HostBirthUtc = ReaderBirth, ParentImage = @"C:\Windows\powershell.exe", HostImage = @"D:\SW\SLDWORKS.exe",
            ReaderImage = @"C:\Windows\powershell.exe", Tests = @"E:\Qualification\VBAi.Tests.dll", TestsSha256 = ReaderHash,
            TestsMvid = "01234567-89ab-cdef-0123-456789abcdef", Macro = Scope().Macro, Receipt = @"E:\Qualification\" + ReaderNonce + ".response.json"
        };
        private static Q020TaskDialogTeardown.ReaderResponse ReaderResponse() => new Q020TaskDialogTeardown.ReaderResponse
        {
            Version = 1, Nonce = ReaderNonce, RequestSha256 = ReaderHash, TestsSha256 = ReaderHash, TestsMvid = ReaderRequest().TestsMvid,
            ReaderPid = 43, ReaderBirthUtc = ReaderBirth, ReaderImage = ReaderRequest().ReaderImage, HostPid = 42, HostBirthUtc = ReaderBirth,
            Completed = true, ReadOnly = true, Success = true, Snapshots = new[] { Snapshot() }
        };

        [DataTestMethod, DataRow(0), DataRow(1), DataRow(2)]
        public void AnExitedIdentityBoundReadOnlyWorkerCanReturnOnlyCompleteSnapshots(int count)
        {
            var request = ReaderRequest(); var result = ReaderResponse();
            Q020TaskDialogTeardown.ValidateReaderRequest(request, ReaderManifest, request.ReaderImage, ReaderHash, request.TestsMvid);
            result.Snapshots = Enumerable.Range(0, count).Select(i => Snapshot(1000 + 1000 * i)).ToArray();
            Assert.AreEqual(count, Q020TaskDialogTeardown.RequireReaderTerminal(request, result, ReaderHash, 43, ReaderBirth, true, 0).Length);
        }

        [DataTestMethod]
        [DataRow("version")][DataRow("nonce")][DataRow("sid")][DataRow("parent")][DataRow("same pid")][DataRow("parent birth")][DataRow("host birth")]
        [DataRow("hash shape")][DataRow("hash case")][DataRow("hash mismatch")][DataRow("mvid")][DataRow("mvid mismatch")]
        [DataRow("host")][DataRow("thread")][DataRow("same root")][DataRow("macro")][DataRow("manifest")][DataRow("image")]
        [DataRow("reader image")][DataRow("parent image")][DataRow("host image")][DataRow("tests")][DataRow("receipt name")][DataRow("receipt directory")][DataRow("unc")]
        public void IncompleteChangedOrForeignReaderRequestsAreRejectedBeforeProviderAcquisition(string fault)
        {
            var r = ReaderRequest(); string manifest = ReaderManifest, image = r.ReaderImage, hash = ReaderHash, mvid = r.TestsMvid;
            if (fault == "version") r.Version = 2; if (fault == "nonce") r.Nonce = "invalid"; if (fault == "sid") r.UserSid = " ";
            if (fault == "parent") r.ParentPid = 0; if (fault == "same pid") r.ParentPid = r.HostPid;
            if (fault == "parent birth") r.ParentBirthUtc = null; if (fault == "host birth") r.HostBirthUtc = ReaderBirth.Replace("Z", "+02:00");
            if (fault == "hash shape") r.TestsSha256 = "wrong"; if (fault == "hash case") r.TestsSha256 = ReaderHash.ToLowerInvariant();
            if (fault == "hash mismatch") hash = new string('B', 64); if (fault == "mvid") r.TestsMvid = "wrong"; if (fault == "mvid mismatch") mvid = Guid.NewGuid().ToString("D");
            if (fault == "host") r.HostPid = 0; if (fault == "thread") r.Thread = 0; if (fault == "same root") r.NativeRoot = r.VbeRoot;
            if (fault == "macro") r.Macro = "Owned.swp"; if (fault == "manifest") manifest = "relative.json";
            if (fault == "image") image = @"C:\Different\powershell.exe"; if (fault == "reader image") { r.ReaderImage = @"C:\Windows\other.exe"; image = r.ReaderImage; }
            if (fault == "parent image") r.ParentImage = @"C:\Windows\other.exe"; if (fault == "host image") r.HostImage = @"D:\SW\Other.exe";
            if (fault == "tests") r.Tests = @"E:\Qualification\Other.dll"; if (fault == "receipt name") r.Receipt = @"E:\Qualification\wrong.json";
            if (fault == "receipt directory") r.Receipt = @"E:\Other\" + ReaderNonce + ".response.json"; if (fault == "unc") r.Tests = @"\\server\share\VBAi.Tests.dll";
            Action validate = () => Q020TaskDialogTeardown.ValidateReaderRequest(r, manifest, image, hash, mvid);
            if (new[] { "host", "thread", "same root", "macro" }.Contains(fault)) Assert.ThrowsException<ArgumentException>(validate);
            else Assert.ThrowsException<InvalidOperationException>(validate);
        }

        [DataTestMethod]
        [DataRow("hash")][DataRow("pid")][DataRow("same parent")][DataRow("birth")][DataRow("exit missing")][DataRow("completed")][DataRow("readonly")][DataRow("input")]
        [DataRow("version")][DataRow("reader pid")][DataRow("reader birth")][DataRow("reader image")][DataRow("nonce")][DataRow("request hash")]
        [DataRow("host pid")][DataRow("host birth")][DataRow("tests hash")][DataRow("tests mvid")][DataRow("exit code")][DataRow("error")][DataRow("error kind")]
        [DataRow("snapshots")][DataRow("partial snapshots")][DataRow("foreign snapshots")]
        public void NoReceiptOrPidAloneCanSubstituteForOriginalObserverExit(string fault)
        {
            var r = ReaderRequest(); var result = ReaderResponse(); string hash = ReaderHash, birth = ReaderBirth; int pid = 43; bool observed = true; uint exit = 0;
            if (fault == "hash") hash = "wrong"; if (fault == "pid") pid = 0; if (fault == "same parent") pid = r.ParentPid;
            if (fault == "birth") birth = "wrong"; if (fault == "exit missing") observed = false; if (fault == "completed") result.Completed = false;
            if (fault == "readonly") result.ReadOnly = false; if (fault == "input") result.NativeInputSent = true; if (fault == "version") result.Version = 2;
            if (fault == "reader pid") result.ReaderPid = 44; if (fault == "reader birth") result.ReaderBirthUtc = ReaderBirth.Replace("13:", "14:");
            if (fault == "reader image") result.ReaderImage = @"C:\Other.exe"; if (fault == "nonce") result.Nonce = Guid.NewGuid().ToString("N");
            if (fault == "request hash") result.RequestSha256 = new string('B', 64); if (fault == "host pid") result.HostPid = 44;
            if (fault == "host birth") result.HostBirthUtc = "wrong"; if (fault == "tests hash") result.TestsSha256 = new string('B', 64);
            if (fault == "tests mvid") result.TestsMvid = Guid.NewGuid().ToString("D"); if (fault == "exit code") exit = 1;
            if (fault == "error") result.Error = "failure"; if (fault == "error kind") result.ErrorKind = "NotReady";
            if (fault == "snapshots") result.Snapshots = null; if (fault == "partial snapshots") result.Snapshots[0].Buttons = null;
            if (fault == "foreign snapshots") result.Snapshots[0].Pid = 44;
            Assert.ThrowsException<InvalidOperationException>(() => Q020TaskDialogTeardown.RequireReaderTerminal(r, result, hash, pid, birth, observed, exit));
        }

        [DataTestMethod, DataRow("NotReady"), DataRow("ObservationUnavailable"), DataRow("Rejected"), DataRow("RuntimeError")]
        public void AnExitedFailedReaderReportsOnlyItsPreciseObservationFailure(string kind)
        {
            var result = ReaderResponse(); result.Success = false; result.Error = "Read incomplete"; result.ErrorKind = kind; result.Snapshots = null;
            uint code = kind == "NotReady" || kind == "ObservationUnavailable" ? 2U : 1U;
            if (kind == "NotReady") Assert.ThrowsException<Q020TaskDialogTeardown.NotReadyException>(() => Q020TaskDialogTeardown.RequireReaderTerminal(ReaderRequest(), result, ReaderHash, 43, ReaderBirth, true, code));
            else if (kind == "ObservationUnavailable") Assert.ThrowsException<Q020TaskDialogTeardown.ObservationUnavailableException>(() => Q020TaskDialogTeardown.RequireReaderTerminal(ReaderRequest(), result, ReaderHash, 43, ReaderBirth, true, code));
            else Assert.ThrowsException<InvalidOperationException>(() => Q020TaskDialogTeardown.RequireReaderTerminal(ReaderRequest(), result, ReaderHash, 43, ReaderBirth, true, code));
            result.Snapshots = new[] { Snapshot() };
            Assert.ThrowsException<InvalidOperationException>(() => Q020TaskDialogTeardown.RequireReaderTerminal(ReaderRequest(), result, ReaderHash, 43, ReaderBirth, true, code));
        }

        [TestMethod]
        public void MissingReaderInputsCannotImplyACompletedRead()
        {
            Assert.ThrowsException<ArgumentNullException>(() => Q020TaskDialogTeardown.ValidateReaderRequest(null, ReaderManifest, ReaderRequest().ReaderImage, ReaderHash, ReaderRequest().TestsMvid));
            Assert.ThrowsException<ArgumentNullException>(() => Q020TaskDialogTeardown.RequireReaderTerminal(null, ReaderResponse(), ReaderHash, 43, ReaderBirth, true, 0));
            Assert.ThrowsException<ArgumentNullException>(() => Q020TaskDialogTeardown.RequireReaderTerminal(ReaderRequest(), null, ReaderHash, 43, ReaderBirth, true, 0));
            Assert.ThrowsException<InvalidOperationException>(() => Q020TaskDialogTeardown.OriginalReaderBirthUtc(0));
        }

        [TestMethod]
        public void HelperCompilesWithTheActualWindowsPowerShellLanguageBoundaryWithoutRunningNativeCode()
        {
            var directory = new DirectoryInfo(Path.GetDirectoryName(typeof(SolidWorksTaskDialogTeardownTests).Assembly.Location));
            while (directory != null && !File.Exists(Path.Combine(directory.FullName, "tools", "tests", "q020-main", "Q020TaskDialogTeardown.cs"))) directory = directory.Parent;
            Assert.IsNotNull(directory);
            using (var provider = new CSharpCodeProvider())
            {
                var parameters = new CompilerParameters { GenerateInMemory = true, CompilerOptions = "/langversion:5" };
                foreach (string reference in new[] { "System.dll", "System.Core.dll", typeof(AutomationElement).Assembly.Location,
                    typeof(ControlType).Assembly.Location, typeof(System.Windows.Point).Assembly.Location }) parameters.ReferencedAssemblies.Add(reference);
                var result = provider.CompileAssemblyFromFile(parameters, Path.Combine(directory.FullName, "tools", "tests", "q020-main", "Q020TaskDialogTeardown.cs"));
                Assert.IsFalse(result.Errors.HasErrors, string.Join(Environment.NewLine, result.Errors.Cast<CompilerError>().Select(e => e.ToString())));
                Assert.IsNotNull(result.CompiledAssembly.GetType("Q020TaskDialogTeardown+DiscardBatch", true));
            }
        }

        [DataTestMethod]
        [DataRow("pid")][DataRow("thread")][DataRow("vbe")][DataRow("host")][DataRow("same roots")]
        [DataRow("null path")][DataRow("relative")][DataRow("unc")][DataRow("traversal")][DataRow("extension")]
        public void ScopeRejectsMissingOrAmbiguousOwnedIdentity(string fault)
        {
            int pid = 42; uint thread = 73; long vbe = 100, host = 200; string path = Scope().Macro;
            if (fault == "pid") pid = 0; if (fault == "thread") thread = 0; if (fault == "vbe") vbe = 0;
            if (fault == "host") host = -1; if (fault == "same roots") host = vbe;
            if (fault == "null path") path = null; if (fault == "relative") path = "Owned.swp";
            if (fault == "unc") path = @"\\server\share\Owned.swp"; if (fault == "traversal") path = @"E:\Qualification\..\Owned.swp";
            if (fault == "extension") path = @"E:\Qualification\Owned.txt";
            Assert.ThrowsException<ArgumentException>(() => new Q020TaskDialogTeardown.Scope(pid, thread, vbe, host, path));
        }

        [DataTestMethod]
        [DataRow(false, false, false)][DataRow(false, false, true)][DataRow(false, true, false)][DataRow(false, true, true)]
        [DataRow(true, false, false)][DataRow(true, false, true)][DataRow(true, true, false)][DataRow(true, true, true)]
        public void ExactObservedCommandSetIsBoundToOneOrTwoOriginalQuestions(bool french, bool buttonType, bool two)
        {
            var snapshots = two ? new[] { Snapshot(1000, french, buttonType), Snapshot(2000, french, buttonType) } : new[] { Snapshot(1000, french, buttonType) };
            snapshots[0].OwnerRoot = 100; // Both captured native and VBE root owners are admitted.
            var bindings = Q020TaskDialogTeardown.Admit(Scope(), snapshots);
            Assert.AreEqual(two ? 2 : 1, bindings.Length); Assert.AreEqual(1002L, bindings[0].No);
            Assert.IsTrue(bindings[0].Prompt.Contains(Scope().Macro));
        }

        [TestMethod]
        public void InactiveFooterDoesNotChangeTheAvailableQuestionButBecomesARefusalWhenVisible()
        {
            var snapshot = Snapshot();
            snapshot.Elements[5].Name = "Hidden footer"; snapshot.Elements[5].Actionable = true; snapshot.Elements[5].Offscreen = true;
            Assert.AreEqual(1002L, Q020TaskDialogTeardown.Admit(Scope(), new[] { snapshot })[0].No);
            snapshot.Elements[5].Offscreen = false;
            Assert.ThrowsException<InvalidOperationException>(() => Q020TaskDialogTeardown.Admit(Scope(), new[] { snapshot }));
            var construction = new Q020TaskDialogTeardown.NotReadyException("Constructing before input");
            Assert.AreEqual("Constructing before input", construction.Message);
        }

        [DataTestMethod]
        [DataRow("null")][DataRow("empty")][DataRow("third modal")][DataRow("null row")][DataRow("dialog")][DataRow("duplicate modal")]
        [DataRow("dialog is root")][DataRow("pid")][DataRow("thread")][DataRow("class")][DataRow("caption")][DataRow("hidden")][DataRow("disabled")][DataRow("owner")]
        [DataRow("nodes null")][DataRow("nodes empty")][DataRow("nodes large")][DataRow("node null")][DataRow("node pid")][DataRow("node type")]
        [DataRow("question absent")][DataRow("question duplicate")][DataRow("question type")][DataRow("question hidden")][DataRow("question disabled")]
        [DataRow("foreign path")][DataRow("basename only")][DataRow("question case")][DataRow("question suffix")]
        [DataRow("command absent")][DataRow("command duplicate")][DataRow("fourth command")][DataRow("command hidden")][DataRow("command disabled")]
        [DataRow("command type")][DataRow("command label")][DataRow("command id")][DataRow("command handle")]
        [DataRow("extra named")][DataRow("extra action")][DataRow("bad icon")]
        [DataRow("buttons null")][DataRow("buttons missing")][DataRow("button null")][DataRow("button handle")][DataRow("button root")]
        [DataRow("button pid")][DataRow("button thread")][DataRow("button class")][DataRow("button id")][DataRow("button hidden")][DataRow("button disabled")]
        [DataRow("duplicate button")][DataRow("caption differs")][DataRow("physical semantic mismatch")][DataRow("button is dialog")]
        [DataRow("button reused between modals")]
        public void EveryIncompleteForeignOrAmbiguousSnapshotFailsBeforeEitherDeliveryRead(string fault)
        {
            for (int boundary = 1; boundary <= 2; boundary++)
            {
                int reads = 0, posts = 0;
                var broken = Broken(fault);
                Assert.ThrowsException<InvalidOperationException>(() => Q020TaskDialogTeardown.Admit(Scope(), broken));
                var plan = new Q020TaskDialogTeardown.DiscardBatch(Scope(), new[] { Snapshot() },
                    () => ++reads == boundary ? broken : new[] { Snapshot() }, () => { }, _ => { posts++; return true; });
                Assert.ThrowsException<InvalidOperationException>(plan.RequestNext);
                Assert.AreEqual(boundary, reads); Assert.AreEqual(0, posts); Assert.AreEqual(0, plan.Settled);
                Assert.ThrowsException<InvalidOperationException>(plan.RequestNext); Assert.AreEqual(boundary, reads);
                Assert.ThrowsException<InvalidOperationException>(() => plan.ObserveCompletion());
            }
        }

        private static Q020TaskDialogTeardown.Snapshot[] Broken(string fault)
        {
            var s = Snapshot(); var rows = new[] { s };
            if (fault == "null") return null; if (fault == "empty") return new Q020TaskDialogTeardown.Snapshot[0];
            if (fault == "third modal") return new[] { s, Snapshot(2000), Snapshot(3000) };
            if (fault == "null row") return new Q020TaskDialogTeardown.Snapshot[] { null };
            if (fault == "dialog") s.Dialog = 0; if (fault == "duplicate modal") return new[] { s, s };
            if (fault == "dialog is root") s.Dialog = 100; if (fault == "pid") s.Pid = 43; if (fault == "thread") s.Thread = 74;
            if (fault == "class") s.Class = "Other"; if (fault == "caption") s.Caption = "Other";
            if (fault == "hidden") s.Visible = false; if (fault == "disabled") s.Enabled = false; if (fault == "owner") s.OwnerRoot = 300;
            if (fault == "nodes null") s.Elements = null; if (fault == "nodes empty") s.Elements = new Q020TaskDialogTeardown.Element[0];
            if (fault == "nodes large") s.Elements = new Q020TaskDialogTeardown.Element[129]; if (fault == "node null") s.Elements[5] = null;
            if (fault == "node pid") s.Elements[5].Pid = 43; if (fault == "node type") s.Elements[5].Type = null;
            if (fault == "question absent") s.Elements[1].Id = "Absent";
            if (fault == "question duplicate") s.Elements = s.Elements.Concat(new[] { s.Elements[1] }).ToArray();
            if (fault == "question type") s.Elements[1].Type = "ControlType.Pane";
            if (fault == "question hidden") s.Elements[1].Offscreen = true; if (fault == "question disabled") s.Elements[1].Enabled = false;
            if (fault == "foreign path") s.Elements[1].Name = s.Elements[1].Name.Replace("Owned", "Other");
            if (fault == "basename only") s.Elements[1].Name = "Enregistrer les modifications apportées à Owned.swp ?";
            if (fault == "question case") s.Elements[1].Name = s.Elements[1].Name.ToUpperInvariant();
            if (fault == "question suffix") s.Elements[1].Name += " Extra";
            if (fault == "command absent") s.Elements[3].Id = "Absent";
            if (fault == "command duplicate") s.Elements[4].Id = s.Elements[3].Id;
            if (fault == "fourth command") s.Elements = s.Elements.Concat(new[] { Command("9", "Help", 1010, false) }).ToArray();
            if (fault == "command hidden") s.Elements[3].Offscreen = true; if (fault == "command disabled") s.Elements[3].Enabled = false;
            if (fault == "command type") s.Elements[3].Type = "ControlType.Text"; if (fault == "command label") s.Elements[3].Name = "Yes";
            if (fault == "command id") s.Elements[3].Id = "CommandButton_8"; if (fault == "command handle") s.Elements[3].Handle = 0;
            if (fault == "extra named") s.Elements[5].Name = "Unexpected"; if (fault == "extra action") s.Elements[5].Actionable = true;
            if (fault == "bad icon") s.Elements[0].Name = "Unexpected";
            if (fault == "buttons null") s.Buttons = null; if (fault == "buttons missing") s.Buttons = s.Buttons.Take(2).ToArray();
            if (fault == "button null") s.Buttons[1] = null; if (fault == "button handle") s.Buttons[1].Handle = 0;
            if (fault == "button root") s.Buttons[1].Root = 2000; if (fault == "button pid") s.Buttons[1].Pid = 43;
            if (fault == "button thread") s.Buttons[1].Thread = 74; if (fault == "button class") s.Buttons[1].Class = "Static";
            if (fault == "button id") s.Buttons[1].Id = 7; if (fault == "button hidden") s.Buttons[1].Visible = false;
            if (fault == "button disabled") s.Buttons[1].Enabled = false; if (fault == "duplicate button") s.Buttons[1].Handle = s.Buttons[0].Handle;
            if (fault == "caption differs") s.Buttons[1].Caption = "&Yes"; if (fault == "physical semantic mismatch") s.Buttons[1].Handle = 9000;
            if (fault == "button is dialog") { s.Buttons[1].Handle = s.Dialog; s.Elements[3].Handle = s.Dialog; }
            if (fault == "button reused between modals") { var other = Snapshot(2000); other.Buttons[1].Handle = s.Buttons[1].Handle; other.Elements[3].Handle = s.Buttons[1].Handle; return new[] { s, other }; }
            return rows;
        }

        [DataTestMethod]
        [DataRow(1)][DataRow(2)][DataRow(3)]
        public void ScopeLossAtEveryPreDeliveryBoundaryLatchesWithoutPostOrRetry(int boundary)
        {
            int checks = 0, posts = 0;
            var plan = new Q020TaskDialogTeardown.DiscardBatch(Scope(), new[] { Snapshot() }, () => new[] { Snapshot() },
                () => { if (++checks == boundary) throw new InvalidOperationException("Scope revoked"); }, _ => { posts++; return true; });
            Assert.ThrowsException<InvalidOperationException>(plan.RequestNext); Assert.AreEqual(0, posts);
            Assert.ThrowsException<InvalidOperationException>(plan.RequestNext); Assert.AreEqual(boundary, checks);
        }

        [DataTestMethod]
        [DataRow(false)][DataRow(true)]
        public void FailedOrThrowingDeliveryCannotBeRetriedOrSettled(bool throws)
        {
            int posts = 0;
            var plan = new Q020TaskDialogTeardown.DiscardBatch(Scope(), new[] { Snapshot() }, () => new[] { Snapshot() }, () => { },
                _ => { posts++; if (throws) throw new InvalidOperationException("Native delivery uncertain"); return false; });
            Assert.ThrowsException<InvalidOperationException>(plan.RequestNext);
            Assert.ThrowsException<InvalidOperationException>(plan.RequestNext);
            Assert.ThrowsException<InvalidOperationException>(() => plan.ObserveCompletion());
            Assert.AreEqual(1, posts); Assert.IsFalse(plan.Complete);
        }

        [TestMethod]
        public void TwoCapturedQuestionsAdvanceOnlyAfterIndependentDisappearanceAndNeverDeliverAgain()
        {
            var current = new[] { Snapshot(), Snapshot(2000) }; var delivered = new System.Collections.Generic.List<long>();
            var plan = new Q020TaskDialogTeardown.DiscardBatch(Scope(), current, () => current, () => { }, handle => { delivered.Add(handle); return true; });
            Assert.ThrowsException<InvalidOperationException>(() => plan.ObserveCompletion());
            plan.RequestNext(); Assert.AreEqual(1000L, plan.PendingDialog); Assert.IsFalse(plan.ObserveCompletion());
            Assert.ThrowsException<InvalidOperationException>(plan.RequestNext); Assert.AreEqual(1, delivered.Count);
            current = new[] { Snapshot(2000) }; Assert.IsTrue(plan.ObserveCompletion()); Assert.AreEqual(1, plan.Settled);
            plan.RequestNext(); Assert.IsFalse(plan.ObserveCompletion());
            current = new Q020TaskDialogTeardown.Snapshot[0]; Assert.IsTrue(plan.ObserveCompletion());
            Assert.IsTrue(plan.Complete); Assert.AreEqual(0L, plan.PendingDialog);
            CollectionAssert.AreEqual(new long[] { 1002, 2002 }, delivered);
            Assert.ThrowsException<InvalidOperationException>(plan.RequestNext);
        }

        [DataTestMethod]
        [DataRow("dialog")][DataRow("no")][DataRow("yes")][DataRow("cancel")][DataRow("owner")][DataRow("order")]
        [DataRow("new after no")][DataRow("null")][DataRow("read throws")][DataRow("scope throws")]
        public void PostDeliveryChangesStayUncertainAndNeverAdvanceOrReplay(string fault)
        {
            bool after = false; int posts = 0;
            var current = new[] { Snapshot(), Snapshot(2000) };
            var plan = new Q020TaskDialogTeardown.DiscardBatch(Scope(), current,
                () => { if (after && fault == "read throws") throw new InvalidOperationException("Read failed"); return current; },
                () => { if (after && fault == "scope throws") throw new InvalidOperationException("Scope revoked"); }, _ => { posts++; return true; });
            plan.RequestNext(); after = true;
            if (fault == "dialog") current[0].Dialog = 3000;
            if (fault == "no") { current[0].Buttons[1].Handle = 9000; current[0].Elements[3].Handle = 9000; }
            if (fault == "yes") { current[0].Buttons[0].Handle = 9000; current[0].Elements[2].Handle = 9000; }
            if (fault == "cancel") { current[0].Buttons[2].Handle = 9000; current[0].Elements[4].Handle = 9000; }
            if (fault == "owner") current[0].OwnerRoot = 100;
            if (fault == "order") current = current.Reverse().ToArray();
            if (fault == "new after no") current = new[] { Snapshot(2000), Snapshot(3000) };
            if (fault == "null") current = null;
            Assert.ThrowsException<InvalidOperationException>(() => plan.ObserveCompletion());
            Assert.AreEqual(0, plan.Settled); Assert.IsFalse(plan.Complete);
            Assert.ThrowsException<InvalidOperationException>(plan.RequestNext); Assert.AreEqual(1, posts);
        }

        [TestMethod]
        public void CapturedBindingsAreImmutableEvenWhenTheCallerMutatesItsApprovedDto()
        {
            var approved = new[] { Snapshot() }; int posts = 0;
            var plan = new Q020TaskDialogTeardown.DiscardBatch(Scope(), approved, () => new[] { Snapshot() }, () => { },
                handle => { Assert.AreEqual(1002L, handle); posts++; return true; });
            approved[0].Dialog = 5000; approved[0].Buttons[1].Handle = 5002; approved[0].Elements[1].Name = "Changed by caller";
            plan.RequestNext(); Assert.AreEqual(1, posts);
        }

        [TestMethod]
        public void MissingDependenciesAndObservationFailuresNeverImplyAnEnteredNativeAction()
        {
            Assert.ThrowsException<ArgumentNullException>(() => Q020TaskDialogTeardown.Admit(null, new[] { Snapshot() }));
            Assert.ThrowsException<ArgumentNullException>(() => new Q020TaskDialogTeardown.DiscardBatch(Scope(), new[] { Snapshot() }, null, () => { }, _ => true));
            Assert.ThrowsException<ArgumentNullException>(() => new Q020TaskDialogTeardown.DiscardBatch(Scope(), new[] { Snapshot() }, () => new[] { Snapshot() }, null, _ => true));
            Assert.ThrowsException<ArgumentNullException>(() => new Q020TaskDialogTeardown.DiscardBatch(Scope(), new[] { Snapshot() }, () => new[] { Snapshot() }, () => { }, null));
            int posts = 0;
            var plan = new Q020TaskDialogTeardown.DiscardBatch(Scope(), new[] { Snapshot() }, () => { throw new InvalidOperationException("Read failed"); }, () => { }, _ => { posts++; return true; });
            Assert.ThrowsException<InvalidOperationException>(plan.RequestNext); Assert.AreEqual(0, posts);
            Assert.ThrowsException<InvalidOperationException>(plan.RequestNext);
        }
    }
}
