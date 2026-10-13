using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using VBAi.Tests.Infrastructure;

namespace VBAi.Tests.Unit.Editor
{
    /// <summary>Checks acquisition ownership and adapter adoption without starting any Office host.</summary>
    [TestClass, TestCategory("Unit"), DoNotParallelize]
    public sealed class EditorModuleOwnershipTests
    {
        private static EditorVbeContract Active()
        {
            var native = new EditorVbeContract();
            native.Vbe.ActiveWindow = native.Original.CodeModule.CodePane.Window;
            return native;
        }

        [STATestMethod, DataRow(false, 5), DataRow(true, 8)]
        public void ActiveProbeOwnsEachGetterButNeverTheBorrowedVbe(bool followOnly, int expected)
        {
            var native = Active();
            var released = new List<object>();
            var module = EditorVbeModule.AcquireActive(native.Vbe, followOnly, released.Add);
            Assert.AreSame(native.Project, module.Project);
            Assert.AreSame(native.Original, module.Component);
            Assert.AreEqual(0, released.Count);
            module.Dispose();
            Assert.AreEqual(expected, released.Count);
            Assert.IsFalse(released.Contains(native.Vbe));
            Assert.AreEqual(followOnly ? 2 : 1, released.Count(item => ReferenceEquals(item, native.Project)));
            Assert.AreEqual(0, native.Original.CodeModule.CodePane.Window.Closes);
            module.Dispose();
            Assert.AreEqual(expected, released.Count, "Disposal must never replay a release.");
        }

        [STATestMethod]
        [DataRow("designer", 1), DataRow("no-window", 0), DataRow("no-pane", 1)]
        [DataRow("running", 6), DataRow("break", 6), DataRow("protected", 6)]
        [DataRow("type", 6), DataRow("foreign-project", 8), DataRow("foreign-window", 8)]
        public void IneligibleActiveProbeReleasesOnlyTheAcquisitionsItReached(string state, int expected)
        {
            var native = Active();
            if (state == "designer") native.Original.CodeModule.CodePane.Window.Type = 1;
            if (state == "no-window") native.Vbe.ActiveWindow = null;
            if (state == "no-pane") native.Vbe.ActiveCodePane = null;
            if (state == "running") native.Project.Mode = 0;
            if (state == "break") native.Project.Mode = 1;
            if (state == "protected") native.Project.Protection = 1;
            if (state == "type") native.Original.Type = 11;
            if (state == "foreign-project") native.Vbe.ActiveVBProject = new object();
            if (state == "foreign-window") native.Vbe.ActiveWindow = new EditorVbeContract.Window { Type = 0 };
            int reads = 0;
            native.Original.CodeModule.BeforeRead = () => reads++;
            var released = new List<object>();
            Assert.IsNull(EditorVbeModule.AcquireActive(native.Vbe, true, released.Add));
            Assert.AreEqual(expected, released.Count, state);
            Assert.AreEqual(0, reads);
            Assert.IsFalse(released.Contains(native.Vbe));
        }

        [STATestMethod]
        public void GetterFailureReleasesTheEarlierAcquisitionsWithoutReadingOrActivating()
        {
            var native = Active();
            native.Original.CodeModule.Parent = null;
            var released = new List<object>();
            Assert.ThrowsException<Microsoft.CSharp.RuntimeBinder.RuntimeBinderException>(() =>
                EditorVbeModule.AcquireActive(native.Vbe, true, released.Add));
            Assert.AreEqual(3, released.Count);
            Assert.AreEqual(0, native.Original.CodeModule.CodePane.Shows);
        }

        [STATestMethod]
        public void WrongThreadDisposalRefusesBeforeReleasingAnyOwnedAcquisition()
        {
            var native = Active(); var released = new List<object>();
            var module = EditorVbeModule.AcquireActive(native.Vbe, true, released.Add);
            Exception failure = null;
            var thread = new Thread(() => { try { module.Dispose(); } catch (Exception error) { failure = error; } });
            thread.Start(); thread.Join();
            Assert.IsInstanceOfType(failure, typeof(InvalidOperationException));
            Assert.AreEqual(0, released.Count);
            module.Dispose(); Assert.AreEqual(8, released.Count);
        }

        [STATestMethod]
        public void FailedReleaseIsRetainedAndDisposalDoesNotRetryIt()
        {
            var native = Active(); int calls = 0;
            var failure = new InvalidOperationException("Owned release boundary failure");
            var module = EditorVbeModule.AcquireActive(native.Vbe, true, value => { calls++; throw failure; });
            Assert.AreSame(failure, Assert.ThrowsException<InvalidOperationException>(() => module.Dispose()));
            Assert.AreEqual(1, calls);
            module.Dispose(); Assert.AreEqual(1, calls);
        }

        [STATestMethod]
        public void BorrowedConstructorDisposalDoesNotReleaseCallerBindingsOrCloseTheirPane()
        {
            var native = Active();
            native.Adapter.Dispose();
            Assert.AreSame(native.Project, native.Vbe.VBProjects.Single());
            Assert.AreSame(native.Original, native.Project.VBComponents.Items.Single());
            Assert.AreEqual(0, native.Original.CodeModule.CodePane.Window.Closes);
        }

        [STATestMethod]
        public void RepeatedRetainedPaneAcquisitionBalancesTheReplacedScopeWithoutNativeClose()
        {
            var native = Active(); var released = new List<object>();
            var module = EditorVbeModule.AcquireActive(native.Vbe, false, released.Add);
            module.RetainNativeWindow();
            released.Clear(); // Ignore Validate's completed temporary scopes.
            module.RetainNativeWindow();
            Assert.AreEqual(1, released.Count(item => ReferenceEquals(item, native.Original.CodeModule.CodePane.Window)));
            Assert.AreEqual(0, native.Original.CodeModule.CodePane.Window.Closes);
            released.Clear(); module.Dispose();
            Assert.AreEqual(8, released.Count, "Five binding-probe getters and three retained-pane getters.");
            Assert.AreEqual(0, native.Original.CodeModule.CodePane.Window.Closes);
        }

        [STATestMethod]
        public void SiblingOwnsItsIndependentComponentAndProjectWhileTemporaryEnumerationIsReleased()
        {
            var native = Active(); var released = new List<object>();
            var root = EditorVbeModule.AcquireActive(native.Vbe, false, released.Add);
            var sibling = root.Sibling(native.Original.Name);
            Assert.IsTrue(released.Count > 0);
            root.Dispose();
            released.Clear(); sibling.Dispose();
            Assert.AreEqual(2, released.Count);
            Assert.AreSame(native.Original, released[0]);
            Assert.AreSame(native.Project, released[1]);
        }

        [STATestMethod]
        public void ReadingReleasesValidationAndCodeGettersBeforeTheOwnedAdapterIsDisposed()
        {
            var native = Active(); var released = new List<object>();
            var module = EditorVbeModule.AcquireActive(native.Vbe, false, released.Add);
            Assert.IsTrue(module.Read().Contains("Public Sub Hello()"));
            Assert.AreEqual(1, released.Count(item => ReferenceEquals(item, native.Original.CodeModule)));
            released.Clear(); module.Dispose(); Assert.AreEqual(5, released.Count);
        }

        [STATestMethod, DataRow(1, true), DataRow(2, true), DataRow(3, false), DataRow(100, true), DataRow(11, false)]
        public void SelectedProbeTransfersOnlyEligibleCodeComponents(int type, bool eligible)
        {
            var native = Active(); native.Original.Type = type;
            var host = new EditorNavigationFixture.Host { Selection = native.Original };
            var released = new List<object>();
            var module = EditorVbeModule.AcquireSelected(host, released.Add);
            if (eligible)
            {
                Assert.AreSame(native.Original, module.Component); Assert.AreEqual(0, released.Count);
                module.Dispose(); Assert.AreEqual(3, released.Count);
            }
            else { Assert.IsNull(module); Assert.AreEqual(1, released.Count); }
            Assert.IsFalse(released.Contains(host));
        }

        [STATestMethod]
        public void OpeningAnExistingDocumentReleasesTheDuplicateAdapterAndKeepsTheAdoptedOne()
        {
            using (var fixture = new ModernEditorDebugFixture())
            {
                var native = fixture.Native;
                var released = new List<object>();
                var duplicate = EditorVbeModule.AcquireActive(native.Vbe, false, released.Add);
                var document = ModernEditorDebugFixture.Wait(fixture.Window.OpenModule(duplicate));
                Assert.AreSame(fixture.Document, document);
                Assert.AreSame(native.Adapter, document.Module);
                Assert.AreEqual(5, released.Count);
                Assert.IsTrue(native.Adapter.Read().Contains("Hello"));
            }
        }

        [STATestMethod]
        public void InitialSourceFailureReleasesTheUnadoptedAdapterAndLeavesItsNativePaneOpen()
        {
            using (var fixture = new ModernEditorDebugFixture())
            {
                var native = Active(); var released = new List<object>();
                var module = EditorVbeModule.AcquireActive(native.Vbe, false, released.Add);
                native.Original.CodeModule.BeforeRead = () => { throw new InvalidOperationException("Owned read failure"); };
                Assert.ThrowsException<InvalidOperationException>(() => ModernEditorDebugFixture.Wait(fixture.Window.OpenModule(module)));
                Assert.IsFalse(fixture.Window.OwnsModule(module));
                Assert.ThrowsException<ObjectDisposedException>(() => module.Read());
                Assert.AreEqual(0, native.Original.CodeModule.CodePane.Window.Closes);
                Assert.IsTrue(released.Contains(native.Project));
            }
        }

        [STATestMethod]
        public void AdoptedDocumentKeepsItsBindingsUntilRuntimeTeardownAfterRendererFailure()
        {
            using (var fixture = new ModernEditorDebugFixture())
            {
                var native = Active(); var released = new List<object>();
                var module = EditorVbeModule.AcquireActive(native.Vbe, false, released.Add);
                fixture.Ready(true);
                fixture.Window.ScriptExecution = (method, values) => method == "open"
                    ? Task.FromException<string>(new InvalidOperationException("Owned renderer failure")) : Task.FromResult("null");
                Assert.ThrowsException<InvalidOperationException>(() => ModernEditorDebugFixture.Wait(fixture.Window.OpenModule(module)));
                Assert.IsTrue(fixture.Window.OwnsModule(module));
                Assert.IsTrue(module.Read().Contains("Hello"));
                released.Clear(); fixture.Window.Dispose();
                Assert.IsTrue(released.Contains(native.Project));
                Assert.ThrowsException<ObjectDisposedException>(() => module.Read());
                Assert.AreEqual(1, native.Original.CodeModule.CodePane.Window.Closes);
            }
        }

        [STATestMethod, DataRow("hidden"), DataRow("same-document"), DataRow("same-transition")]
        public void PassiveObservationReleasesEveryUnusedProbeWithoutClosingTheNativePane(string state)
        {
            using (var fixture = new ModernEditorDebugFixture())
            {
                var native = fixture.Native;
                native.Vbe.ActiveWindow = native.Original.CodeModule.CodePane.Window;
                var released = new List<object>(); int probes = 0;
                fixture.Window.WorkspaceHosted = true; fixture.Ready(true);
                fixture.Window.Show();
                fixture.Window.ReadActiveModule = () => { probes++; return EditorVbeModule.AcquireActive(native.Vbe, true, released.Add); };
                if (state == "hidden") fixture.Window.Hide();
                if (state == "same-transition") ModernEditorDebugFixture.Wait(fixture.Window.FollowNativeActivation());
                released.Clear(); probes = 0;
                ModernEditorDebugFixture.Wait(fixture.Window.FollowNativeActivation());
                Assert.AreEqual(1, probes);
                Assert.AreEqual(8, released.Count);
                Assert.AreSame(native.Adapter, fixture.Document.Module);
                Assert.AreEqual(0, native.Original.CodeModule.CodePane.Window.Closes);
            }
        }

        [STATestMethod, DataRow("capture-failure"), DataRow("changed-selection"), DataRow("closing")]
        public void PendingCaptureCannotLeakAnAbandonedTargetOrItsRecheckProbe(string state)
        {
            using (var fixture = new ModernEditorDebugFixture())
            {
                var native = Active(); var released = new List<object>();
                var capture = new TaskCompletionSource<string>();
                fixture.Window.WorkspaceHosted = true; fixture.Ready(true); fixture.Window.Show();
                fixture.Window.ScriptExecution = (method, values) => method == "snapshots" ? capture.Task : Task.FromResult("1");
                fixture.Window.ReadActiveModule = () => EditorVbeModule.AcquireActive(native.Vbe, true, released.Add);
                var follow = fixture.Window.FollowNativeActivation();
                System.Windows.Forms.Application.DoEvents();
                Assert.IsFalse(follow.IsCompleted);
                if (state == "changed-selection") native.Vbe.ActiveCodePane = null;
                if (state == "closing") fixture.Set("closing", true);
                if (state == "capture-failure") capture.SetException(new InvalidOperationException("Owned capture failure"));
                else capture.SetResult("null");
                ModernEditorDebugFixture.Wait(follow);
                fixture.Set("closing", false);
                Assert.IsFalse(fixture.Window.Documents.Any(document => document.Module is EditorVbeModule adapter && ReferenceEquals(adapter.Project, native.Project)));
                Assert.AreEqual(state == "changed-selection" ? 9 : 8, released.Count);
                Assert.AreEqual(0, native.Original.CodeModule.CodePane.Window.Closes);
            }
        }

        [STATestMethod]
        public void SuccessfulPassiveObservationAdoptsTheTargetAndReleasesOnlyTheRecheckProbe()
        {
            using (var fixture = new ModernEditorDebugFixture())
            {
                var native = Active(); var released = new List<object>();
                var probes = new List<EditorVbeModule>();
                fixture.Window.WorkspaceHosted = true; fixture.Ready(true); fixture.Window.Show();
                fixture.Window.ScriptExecution = (method, values) => Task.FromResult(method == "snapshots" ? "null" : "1");
                fixture.Window.ReadActiveModule = () =>
                {
                    var module = EditorVbeModule.AcquireActive(native.Vbe, true, released.Add);
                    probes.Add(module); return module;
                };
                ModernEditorDebugFixture.Wait(fixture.Window.FollowNativeActivation());
                Assert.AreEqual(2, probes.Count);
                Assert.IsTrue(fixture.Window.OwnsModule(probes[0]));
                Assert.IsFalse(fixture.Window.OwnsModule(probes[1]));
                Assert.ThrowsException<ObjectDisposedException>(() => probes[1].Read());
                Assert.IsTrue(probes[0].Read().Contains("Hello"));
                int count = released.Count;
                fixture.Window.Dispose();
                Assert.IsTrue(released.Count > count);
                Assert.ThrowsException<ObjectDisposedException>(() => probes[0].Read());
            }
        }

        [STATestMethod, DataRow("success"), DataRow("prepared"), DataRow("renamed"), DataRow("removed")]
        public void AttributeReplacementTracksImportedBindingsAcrossCommitAndRecovery(string checkpoint)
        {
            var native = Active();
            native.Original.CodeModule.Raw = "Public Sub Hello( _\n    Optional ByVal count As Long = 1)\nAttribute Hello.VB_Description = \"Keep\"\n    Debug.Print count\nEnd Sub";
            var released = new List<object>();
            var module = EditorVbeModule.AcquireActive(native.Vbe, false, released.Add);
            string before = EditorDocument.Normalize(module.Read());
            string after = before.Replace("count As Long = 1", "count As Long = 2");
            module.AttributeRewriteCheckpoint = boundary =>
            {
                if (boundary == checkpoint) throw new InvalidOperationException("Owned rewrite interruption");
            };
            if (checkpoint == "success") Assert.AreEqual(after, EditorDocument.Normalize(module.Write(before, after)));
            else Assert.ThrowsException<InvalidOperationException>(() => module.Write(before, after));
            Assert.AreEqual(checkpoint == "success" ? after : before, EditorDocument.Normalize(module.Read()));
            object current = module.Component;
            released.Clear(); module.Dispose();
            Assert.IsTrue(released.Contains(current), "The restored or committed component has its own getter acquisition.");
            Assert.AreEqual(1, native.Project.VBComponents.Items.Count);
        }

        [STATestMethod, DataRow("capacity"), DataRow("invalid-key"), DataRow("missing-pane-window")]
        public void PreAdoptionFailureCannotRetainTheIncomingAdapter(string boundary)
        {
            using (var fixture = new ModernEditorDebugFixture())
            {
                var native = Active(); var released = new List<object>();
                var module = EditorVbeModule.AcquireActive(native.Vbe, false, released.Add);
                var originalWindow = native.Original.CodeModule.CodePane.Window;
                if (boundary == "capacity")
                {
                    var documents = fixture.Get<Dictionary<string, EditorDocument>>("documents");
                    while (documents.Count < 30)
                    {
                        var extra = new EditorDocument(fixture.Storage); documents.Add(extra.Id, extra);
                    }
                    Assert.ThrowsException<InvalidOperationException>(() => ModernEditorDebugFixture.Wait(fixture.Window.OpenModule(module)));
                }
                else if (boundary == "invalid-key")
                {
                    native.Project.Path = "C:\\invalid\0.bas";
                    Assert.ThrowsException<ArgumentException>(() => ModernEditorDebugFixture.Wait(fixture.Window.OpenModule(module)));
                }
                else
                {
                    LlmBoundaryScope.Set(native.Original.CodeModule.CodePane, "<Window>k__BackingField", null);
                    Assert.ThrowsException<InvalidOperationException>(() => ModernEditorDebugFixture.Wait(fixture.Window.OpenModule(module)));
                }
                Assert.IsFalse(fixture.Window.OwnsModule(module));
                if (boundary != "missing-pane-window") Assert.AreEqual(5, released.Count);
                Assert.IsTrue(released.Contains(native.Project));
                Assert.AreEqual(0, originalWindow.Closes);
            }
        }

        [STATestMethod]
        public void ClosingOneTabSavesItsDraftAndReleasesOnlyItsAdoptedAdapter()
        {
            using (var fixture = new ModernEditorDebugFixture())
            using (var dispatcher = new OwnedEditorDispatcher())
            {
                var native = Active(); var released = new List<object>();
                var module = EditorVbeModule.AcquireActive(native.Vbe, false, released.Add);
                var document = ModernEditorDebugFixture.Wait(fixture.Window.OpenModule(module));
                document.Edit(document.Text + "\n' preserved ownership draft");
                fixture.Ready(true); released.Clear();
                UiInvoke.Call(typeof(ModernEditorWindow), "CloseModuleClick", fixture.Window, null, EventArgs.Empty);
                dispatcher.Drain();
                Assert.IsFalse(fixture.Window.OwnsModule(module));
                Assert.IsTrue(fixture.Window.OwnsModule(fixture.Native.Adapter));
                Assert.IsTrue(fixture.Native.Adapter.Read().Contains("Hello"));
                Assert.AreEqual(document.Text, fixture.Window.Drafts.Recover(document.RecoveryKey).Text);
                Assert.AreEqual(1, native.Original.CodeModule.CodePane.Window.Closes);
                Assert.ThrowsException<ObjectDisposedException>(() => module.Read());
                Assert.IsTrue(released.Contains(native.Project));
            }
        }

        [STATestMethod, DataRow("disposed-before-delivery"), DataRow("delivery-failure"), DataRow("queue-failure")]
        public void QueuedNavigationReleasesAdaptersThatWereNeverDelivered(string boundary)
        {
            using (var fixture = new EditorNavigationFixture())
            {
                var native = Active(); fixture.Vbe.Selection = native.Original;
                fixture.Vbe.VBProjects.Add(native.Project);
                int opens = 0;
                var navigation = new EditorProjectNavigation(fixture.Vbe, fixture.Dispatcher, module =>
                {
                    opens++; throw new InvalidOperationException("Owned delivery failure");
                });
                try
                {
                    if (boundary == "queue-failure") fixture.Dispatcher.Dispose();
                    bool queued = (bool)EditorNavigationFixture.Invoke(navigation, "Route");
                    if (boundary == "queue-failure") { Assert.IsFalse(queued); Assert.AreEqual(0, opens); }
                    else
                    {
                        var pending = (List<EditorVbeModule>)EditorNavigationFixture.Get(navigation, "pending");
                        Assert.AreEqual(1, pending.Count);
                        var adapter = pending[0];
                        if (boundary == "disposed-before-delivery") navigation.Dispose();
                        System.Windows.Forms.Application.DoEvents();
                        Assert.AreEqual(boundary == "delivery-failure" ? 1 : 0, opens);
                        Assert.AreEqual(0, pending.Count);
                        Assert.ThrowsException<ObjectDisposedException>(() => adapter.Read());
                    }
                }
                finally { navigation.Dispose(); }
                Assert.AreEqual(0, native.Original.CodeModule.CodePane.Window.Closes);
            }
        }

        [STATestMethod, DataRow("plain"), DataRow("attributed"), DataRow("replacement")]
        public void FailedRecoveryRetainsReferencesAndCannotReplayTheMutation(string shape)
        {
            var native = Active(); var released = new List<object>();
            if (shape == "attributed") native.Original.CodeModule.Raw = "Public Sub Hello()\nAttribute Hello.VB_Description = \"Keep\"\n    Debug.Print 1\nEnd Sub";
            if (shape == "replacement") native.Original.CodeModule.Raw = "Public Sub Hello( _\n    Optional ByVal count As Long = 1)\nAttribute Hello.VB_Description = \"Keep\"\n    Debug.Print count\nEnd Sub";
            var module = EditorVbeModule.AcquireActive(native.Vbe, false, released.Add);
            var document = new EditorDocument(module);
            string before = document.Text;
            string after = shape == "replacement" ? before.Replace("count As Long = 1", "count As Long = 2") : before.Replace("Public Sub Hello()", "Public Sub Changed()").Replace("Print 1", "Print 2");
            int mutations = 0;
            native.Original.CodeModule.BeforeOperation = operation =>
            {
                mutations++;
                if ((shape == "plain" && operation == "insert") || (shape == "attributed" && operation == "add"))
                    throw new InvalidOperationException("Owned write and restoration failure");
            };
            if (shape == "replacement")
            {
                native.Project.VBComponents.AfterRemove = removed => { mutations++; native.Project.VBComponents.ImportThrowsBefore = true; };
                module.AttributeRewriteCheckpoint = phase => { if (phase == "removed") throw new InvalidOperationException("Owned replacement recovery failure"); };
            }
            document.Edit(after);
            // The original uncertain result is retained; the second operation must refuse before a mutation.
            try { document.Synchronize(); Assert.Fail("Both write and recovery must fail in this prepared scenario."); }
            catch (Exception failure) { Assert.IsFalse(failure is AssertFailedException); }
            Assert.IsNotNull(module.MutationUnavailability);
            Assert.IsTrue(mutations > 0, "This scenario must reach the owned native mutation before recovery fails.");
            int count = mutations;
            Assert.ThrowsException<InvalidOperationException>(() => module.Write(before, after));
            Assert.AreEqual(count, mutations);
            Assert.IsTrue(document.Dirty);
            released.Clear(); module.Dispose();
            Assert.AreEqual(0, released.Count, "Uncertain-operation acquisitions remain rooted for recovery.");
            Assert.AreEqual(0, native.Original.CodeModule.CodePane.Window.Closes);
        }

        [STATestMethod, DataRow("alias"), DataRow("missing-project"), DataRow("missing-component")]
        public void BindingTransferBalancesAliasAcquisitionsAndRejectsUnownedArguments(string state)
        {
            var released = new List<object>();
            object project = new object(); object component = state == "alias" ? project : new object();
            using (var producer = new VbaTestComReferences(released.Add))
            {
                if (state != "missing-project") producer.Add(project);
                if (state != "missing-component") producer.Add(component);
                if (state == "alias")
                {
                    using (var module = EditorVbeModule.TransferBindings(new object(), project, component, producer, released.Add))
                    {
                        producer.Dispose(); Assert.AreEqual(0, released.Count);
                    }
                    Assert.AreEqual(2, released.Count);
                    Assert.IsTrue(released.All(value => ReferenceEquals(value, project)));
                }
                else Assert.ThrowsException<InvalidOperationException>(() =>
                    EditorVbeModule.TransferBindings(new object(), project, component, producer, released.Add));
            }
            if (state == "missing-project") CollectionAssert.AreEqual(new[] { component }, released);
            if (state == "missing-component") CollectionAssert.AreEqual(new[] { project }, released);
        }

        [STATestMethod, DataRow("native-close"), DataRow("renderer")]
        public void RuntimeFailureStillReleasesEveryAdoptedAdapterAndRetainsTheOriginalFailure(string boundary)
        {
            using (var fixture = new ModernEditorBrowserFixture())
            {
                ModernEditorDebugFixture.Wait(fixture.Initialize());
                var native = Active(); var released = new List<object>();
                var module = EditorVbeModule.AcquireActive(native.Vbe, false, released.Add);
                ModernEditorDebugFixture.Wait(fixture.Window.OpenModule(module));
                var failure = new InvalidOperationException("Owned runtime cleanup failure");
                int closes = 0;
                if (boundary == "native-close") fixture.Editor.Base.Native.Original.CodeModule.CodePane.Window.OnClose = () => { throw failure; };
                else fixture.Controller.BeforeCall = operation => { if (operation == "Close") { closes++; throw failure; } };
                released.Clear();
                Assert.AreSame(failure, Assert.ThrowsException<InvalidOperationException>(() => fixture.Window.Dispose()));
                Assert.IsTrue(released.Contains(native.Project));
                Assert.ThrowsException<ObjectDisposedException>(() => module.Read());
                Assert.AreEqual(1, native.Original.CodeModule.CodePane.Window.Closes);
                fixture.Window.Dispose();
                if (boundary == "renderer") Assert.AreEqual(1, closes, "Renderer teardown must not replay a failed native close.");
            }
        }
    }
}
