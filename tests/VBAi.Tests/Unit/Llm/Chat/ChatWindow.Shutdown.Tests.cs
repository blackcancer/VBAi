using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using VBAi.Tests.Infrastructure;

namespace VBAi.Tests.Unit
{
    /// <summary>Tests bounded observations against actual managed form and HTTP-resource disposal, without a native host.</summary>
    [TestClass, TestCategory("Unit"), DoNotParallelize]
    public sealed class ChatShutdownDiagnosticTests
    {
        private const BindingFlags Private = BindingFlags.Instance | BindingFlags.NonPublic;
        private static readonly string[] Stages = {
            "chat.Runtime.contextMonitor.Stop", "chat.Runtime.referenceEvents.Dispose",
            "chat.Runtime.projectEvents.Dispose", "chat.Runtime.componentEvents.Dispose",
            "chat.Runtime.SaveCurrentSession", "chat.Runtime.streamRender.Stop",
            "chat.Runtime.pendingFollow.Abort", "chat.Runtime.saveTimer.Stop",
            "chat.Runtime.historySearch.Stop", "chat.Runtime.projectRetry.Stop",
            "chat.Runtime.scopeIdentities.Dispose", "chat.Runtime.persistenceWorker.Dispose",
            "chat.Runtime.sessionStore.Dispose", "chat.Runtime.httpClient.Dispose",
            "chat.Runtime.codex.Dispose", "chat.Runtime.changesMenu.Dispose",
            "chat.Runtime.menuImages.Dispose", "chat.Runtime.entryViews.Dispose",
            "chat.Runtime.composer.Dispose", "chat.Components.Dispose", "chat.Base.Dispose"
        };

        private sealed class Observation
        {
            internal readonly List<Dictionary<string, object>> Events = new List<Dictionary<string, object>>();
            internal Action<Dictionary<string, object>> BeforePublish;
            internal readonly AddInShutdownDiagnostic Root;
            internal Observation()
            {
                const string nonce = "e0dc627ec00147c5b2304c503c144afa";
                var json = new JavaScriptSerializer();
                var identity = new AddInShutdownDiagnostic.Identity {
                    ProcessId = 71, ProcessStartedUtc = "2026-10-05T17:27:03.6593947Z",
                    HostImagePath = @"C:\Qualification\EXCEL.EXE", ProductPath = @"C:\Qualification\VBAi.dll",
                    ProductMvid = "f855d463-0d5e-48f3-8c45-1a35db5a1025", ProductSha256 = new string('A', 64), ThreadId = 73
                };
                Root = AddInShutdownDiagnostic.Begin(json.Serialize(new { Version = 1, Nonce = nonce, Identity = identity }),
                    nonce, "Dispose", () => identity,
                    () => new AddInShutdownDiagnostic.ThreadIdentity {
                        ManagedThreadId = Thread.CurrentThread.ManagedThreadId, NativeThreadId = 73, Apartment = "STA"
                    }, (sequence, data, terminal) => {
                        var row = (Dictionary<string, object>)json.DeserializeObject(json.Serialize(data));
                        BeforePublish?.Invoke(row);
                        Events.Add(row);
                    });
            }
        }

        private sealed class FailingHandler : HttpMessageHandler
        {
            internal int DisposeEntries;
            internal readonly IOException Failure = new IOException("Owned transport disposal failure");
            protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
                => throw new AssertFailedException("No HTTP request is authorized by these disposal tests.");
            protected override void Dispose(bool disposing)
            {
                if (!disposing) return;
                DisposeEntries++;
                base.Dispose(disposing);
                throw Failure;
            }
        }

        private sealed class CallbackContainer : IContainer
        {
            internal readonly IContainer Original;
            internal Action OnDispose;
            internal Exception Failure;
            internal int Calls;
            internal CallbackContainer(IContainer original) { Original = original; }
            public ComponentCollection Components => Original.Components;
            public void Add(IComponent component) => Original.Add(component);
            public void Add(IComponent component, string name) => Original.Add(component, name);
            public void Remove(IComponent component) => Original.Remove(component);
            public void Dispose()
            {
                Calls++;
                var callback = OnDispose; OnDispose = null; callback?.Invoke();
                if (Failure != null) throw Failure;
                Original.Dispose();
            }
        }

        private static object Get(ChatWindow window, string field) => typeof(ChatWindow).GetField(field, Private).GetValue(window);
        private static void Set(ChatWindow window, string field, object value) => typeof(ChatWindow).GetField(field, Private).SetValue(window, value);
        private static bool Is(Dictionary<string, object> row, string phase, string stage)
            => (string)row["Phase"] == phase && (string)row["Stage"] == stage;

        [STATestMethod, DataRow(false), DataRow(true)]
        public void ObservedShutdownRetiresActualFormAndComponentsInOrderWithoutCreatingAnUnrequestedHandle(bool handle)
        {
            var observation = new Observation();
            using (var window = new ChatWindow())
            {
                if (handle) { var ignored = window.Handle; }
                int created = 0, disposed = 0;
                window.HandleCreated += (sender, args) => created++;
                window.Disposed += (sender, args) => disposed++;
                var components = (IContainer)Get(window, "components");
                var owned = new Component(); components.Add(owned);
                int componentDisposed = 0; owned.Disposed += (sender, args) => componentDisposed++;
                window.DisposeForHostShutdown(observation.Root);
                Assert.IsTrue(window.IsDisposed);
                Assert.AreEqual(1, disposed); Assert.AreEqual(1, componentDisposed); Assert.AreEqual(0, created);
                Assert.IsNull(Get(window, "shutdownDiagnostic")); Assert.IsFalse(observation.Root.DiagnosticFailed);
                var rootEvents = observation.Events.Where(row => (string)row["InvocationId"] == observation.Root.InvocationId);
                CollectionAssert.AreEqual(Stages, rootEvents.Where(row => (string)row["Phase"] == "Entry" && row["Stage"] != null).Select(row => (string)row["Stage"]).ToArray());
                CollectionAssert.AreEqual(Stages, rootEvents.Where(row => row["Stage"] != null && (string)row["Phase"] == "Returned").Select(row => (string)row["Stage"]).ToArray());
                Assert.IsTrue(Is(observation.Events.Last(), "Returned", null));
                Assert.IsFalse(observation.Events.Any(row => (string)row["Phase"] == "Fault"));
            }
        }

        public static IEnumerable<object[]> PublicationFailures()
        {
            foreach (string stage in Stages)
                foreach (string phase in new[] { "Entry", "Returned" }) yield return new object[] { stage, phase };
            yield return new object[] { null, "Returned" };
        }

        [STATestMethod, DynamicData(nameof(PublicationFailures), DynamicDataSourceType.Method)]
        public void EveryPublicationFailureLeavesActualFormDisposalAndOwnedComponentReleaseIntact(string stage, string phase)
        {
            var observation = new Observation();
            int attempts = 0;
            var error = new IOException("One bounded observer publication failed");
            observation.BeforePublish = row => { if (Is(row, phase, stage)) { attempts++; throw error; } };
            using (var window = new ChatWindow())
            {
                int disposed = 0; window.Disposed += (sender, args) => disposed++;
                var component = new Component(); ((IContainer)Get(window, "components")).Add(component);
                int componentDisposed = 0; component.Disposed += (sender, args) => componentDisposed++;
                window.DisposeForHostShutdown(observation.Root);
                Assert.AreEqual(1, attempts); Assert.AreEqual(1, disposed); Assert.AreEqual(1, componentDisposed);
                Assert.IsTrue(window.IsDisposed); Assert.IsNull(Get(window, "shutdownDiagnostic"));
                Assert.IsTrue(observation.Root.DiagnosticFailed); Assert.AreSame(error, observation.Root.PublicationError);
            }
        }

        [STATestMethod, DataRow(false), DataRow(true)]
        public void HandledTransportDisposalFailureIsNotRetriedAndDoesNotPreventTheFormFromClosing(bool rejectFaultReceipt)
        {
            var observation = new Observation();
            int rejected = 0;
            if (rejectFaultReceipt) observation.BeforePublish = row => {
                if (Is(row, "Fault", "chat.Runtime.httpClient.Dispose")) { rejected++; throw new IOException("Observer failure"); }
            };
            var handler = new FailingHandler();
            using (var window = new ChatWindow())
            {
                var client = new LlmChatClient(LlmProvider.All.Single(p => p.Name == "Ollama"), new LlmSettings(), "fixture-model", handler);
                Set(window, "activeHttpClient", client);
                window.DisposeForHostShutdown(observation.Root);
                Assert.AreEqual(1, handler.DisposeEntries); Assert.IsTrue(window.IsDisposed);
                Assert.IsNull(Get(window, "activeHttpClient")); Assert.IsNull(Get(window, "shutdownDiagnostic"));
                if (rejectFaultReceipt) { Assert.AreEqual(1, rejected); Assert.IsTrue(observation.Root.DiagnosticFailed); }
                else {
                    var fault = observation.Events.Single(row => Is(row, "Fault", "chat.Runtime.httpClient.Dispose"));
                    Assert.AreEqual(true, fault["FaultCaught"]); StringAssert.Contains((string)fault["Error"], handler.Failure.Message);
                    Assert.IsFalse(observation.Events.Any(row => Is(row, "Returned", "chat.Runtime.httpClient.Dispose")));
                    Assert.IsTrue(Is(observation.Events.Last(), "Returned", null)); Assert.IsFalse(observation.Root.DiagnosticFailed);
                }
                client.Dispose(); Assert.AreEqual(1, handler.DisposeEntries, "The original failed transport operation cannot be replayed.");
            }
        }

        [STATestMethod]
        public void ComponentDisposalFailurePropagatesTheSameErrorWithoutClaimingBaseDisposalOrRepeatingTheFailure()
        {
            var observation = new Observation();
            using (var window = new ChatWindow())
            {
                var original = (IContainer)Get(window, "components");
                var failure = new IOException("Unexpected owned component failure");
                var container = new CallbackContainer(original) { Failure = failure }; Set(window, "components", container);
                try {
                    Assert.AreSame(failure, Assert.ThrowsException<IOException>(() => window.DisposeForHostShutdown(observation.Root)));
                    Assert.AreEqual(1, container.Calls); Assert.IsFalse(window.IsDisposed); Assert.IsNull(Get(window, "shutdownDiagnostic"));
                    Assert.IsFalse(observation.Root.DiagnosticFailed);
                    Assert.AreEqual(1, observation.Events.Count(row => Is(row, "Fault", "chat.Components.Dispose")));
                    Assert.IsTrue(Is(observation.Events.Last(), "Fault", null));
                    Assert.IsFalse(observation.Events.Any(row => (string)row["Stage"] == "chat.Base.Dispose"));
                } finally { Set(window, "components", original); }
                Assert.AreEqual(1, container.Calls);
            }
        }

        [STATestMethod]
        public void ReentrantDisposalRestoresTheOuterObserverWithoutRepeatingRuntimeCleanup()
        {
            var observation = new Observation();
            using (var window = new ChatWindow())
            {
                var original = (IContainer)Get(window, "components");
                var container = new CallbackContainer(original); Set(window, "components", container);
                container.OnDispose = () => window.DisposeForHostShutdown(observation.Root.Child("Dispose"));
                try {
                    window.DisposeForHostShutdown(observation.Root);
                    Assert.IsTrue(window.IsDisposed); Assert.AreEqual(2, container.Calls);
                    Assert.IsNull(Get(window, "shutdownDiagnostic")); Assert.IsFalse(observation.Root.DiagnosticFailed);
                    var starts = observation.Events.Where(row => Is(row, "Entry", null)).ToArray();
                    Assert.AreEqual(2, starts.Length); Assert.AreEqual(starts[0]["InvocationId"], starts[1]["ParentInvocationId"]);
                    Assert.AreEqual(19, observation.Events.Count(row => (string)row["Phase"] == "Entry" && ((string)row["Stage"] ?? "").StartsWith("chat.Runtime.", StringComparison.Ordinal)));
                    Assert.IsTrue(Is(observation.Events.Last(), "Returned", null));
                } finally { Set(window, "components", original); }
            }
        }

        [STATestMethod]
        public void DisabledShutdownObservationPreservesOrdinaryFormDisposal()
        {
            using (var window = new ChatWindow())
            {
                int disposed = 0; window.Disposed += (sender, args) => disposed++;
                window.DisposeForHostShutdown(null);
                Assert.IsTrue(window.IsDisposed); Assert.AreEqual(1, disposed); Assert.IsNull(Get(window, "shutdownDiagnostic"));
            }
        }

        [STATestMethod, DataRow(false), DataRow(true)]
        public void EmbeddedShutdownObservesOnlyExistingFrameworkBoundariesWithExactNesting(bool handle)
        {
            var observation = new Observation();
            using (var parent = new System.Windows.Forms.Panel())
            using (var window = new ChatWindow())
            {
                window.TopLevel = false; parent.Controls.Add(window);
                if (handle) { var ignored = window.Handle; }
                int created = 0, removed = 0;
                window.HandleCreated += (sender, args) => created++;
                parent.ControlRemoved += (sender, args) => removed++;
                window.DisposeForHostShutdown(observation.Root);
                Assert.IsTrue(window.IsDisposed); Assert.IsFalse(parent.IsDisposed);
                Assert.IsNull(window.Parent); Assert.AreEqual(0, parent.Controls.Count);
                Assert.AreEqual(0, created); Assert.AreEqual(1, removed);
                Assert.IsFalse(observation.Root.DiagnosticFailed); Assert.IsNull(Get(window, "shutdownDiagnostic"));
                var boundaries = observation.Events.Where(row => (string)row["Phase"] == "Entry" && ((string)row["Stage"] ?? "").StartsWith("chat.NativeHandle.", StringComparison.Ordinal)).ToArray();
                Assert.AreEqual(handle ? 2 : 0, boundaries.Length);
                foreach (var boundary in boundaries)
                    Assert.AreEqual(1, observation.Events.Count(row => Is(row, "Returned", (string)boundary["Stage"]) && Equals(row["InvocationId"], boundary["InvocationId"])));
                var changed = observation.Events.Single(row => Is(row, "Entry", "chat.Parent.ChangedNotification"));
                Assert.AreEqual(observation.Root.InvocationId, changed["ParentInvocationId"]);
                if (handle)
                {
                    var destroy = boundaries.Single(row => (string)row["Stage"] == "chat.NativeHandle.Destroy");
                    var notification = boundaries.Single(row => (string)row["Stage"] == "chat.NativeHandle.DestroyedNotification");
                    Assert.AreEqual(observation.Root.InvocationId, destroy["ParentInvocationId"]);
                    Assert.AreEqual(destroy["InvocationId"], notification["ParentInvocationId"]);
                }
                Assert.IsTrue(Is(observation.Events.Last(), "Returned", null));
            }
        }

        [STATestMethod]
        [DataRow("chat.NativeHandle.Destroy", "Entry")]
        [DataRow("chat.NativeHandle.Destroy", "Returned")]
        [DataRow("chat.NativeHandle.DestroyedNotification", "Entry")]
        [DataRow("chat.NativeHandle.DestroyedNotification", "Returned")]
        [DataRow("chat.Parent.ChangedNotification", "Entry")]
        [DataRow("chat.Parent.ChangedNotification", "Returned")]
        public void FrameworkBoundaryPublicationFailureDoesNotChangeActualEmbeddedDisposal(string stage, string phase)
        {
            var observation = new Observation(); int rejected = 0;
            observation.BeforePublish = row => { if (Is(row, phase, stage)) { rejected++; throw new IOException("Bounded observer failure"); } };
            using (var parent = new System.Windows.Forms.Panel())
            using (var window = new ChatWindow())
            {
                window.TopLevel = false; parent.Controls.Add(window); var ignored = window.Handle;
                int removed = 0; parent.ControlRemoved += (sender, args) => removed++;
                window.DisposeForHostShutdown(observation.Root);
                Assert.IsTrue(window.IsDisposed); Assert.IsFalse(parent.IsDisposed);
                Assert.AreEqual(1, removed); Assert.AreEqual(1, rejected);
                Assert.IsTrue(observation.Root.DiagnosticFailed); Assert.IsNull(Get(window, "shutdownDiagnostic"));
            }
        }

        [STATestMethod]
        public void FrameworkParentNotificationFailurePropagatesUnchangedAndRestoresTheOriginalObserverWithoutRetry()
        {
            var observation = new Observation();
            using (var parent = new System.Windows.Forms.Panel())
            using (var window = new ChatWindow())
            {
                // Parent removal invokes this notification directly. HandleDestroyed is
                // dispatched through NativeWindow, which routes exceptions to a modal
                // ThreadException dialog instead of propagating them to the caller.
                window.TopLevel = false; parent.Controls.Add(window); int calls = 0;
                var failure = new IOException("Original parent notification failure");
                EventHandler subscriber = (sender, args) => { calls++; throw failure; };
                window.ParentChanged += subscriber;
                try
                {
                    Assert.AreSame(failure, Assert.ThrowsException<IOException>(() => window.DisposeForHostShutdown(observation.Root)));
                    Assert.AreEqual(1, calls); Assert.IsNull(Get(window, "shutdownDiagnostic"));
                    Assert.IsFalse(observation.Root.DiagnosticFailed);
                    Assert.IsFalse(observation.Events.Any(row => Is(row, "Returned", "chat.Parent.ChangedNotification")));
                    Assert.AreEqual(1, observation.Events.Count(row => Is(row, "Fault", "chat.Parent.ChangedNotification")));
                    Assert.IsTrue(Is(observation.Events.Last(), "Fault", null));
                }
                finally { window.ParentChanged -= subscriber; }
            }
        }

        [STATestMethod, DataRow(false), DataRow(true)]
        public void CaughtFirstChanceExceptionIsObservedOnlyDuringItsExactShutdownScope(bool com)
        {
            var observation = new Observation();
            using (var window = new ChatWindow())
            {
                var original = (IContainer)Get(window, "components");
                var container = new CallbackContainer(original); Set(window, "components", container);
                Exception failure = com ? (Exception)new InvalidComObjectException("Owned callback test") : new IOException("Owned callback test");
                container.OnDispose = () => { try { throw failure; } catch (Exception) { } };
                window.DisposeForHostShutdown(observation.Root);
                Assert.IsTrue(window.IsDisposed); Assert.IsFalse(observation.Root.DiagnosticFailed);
                var row = observation.Events.Single(item => Is(item, "ExceptionObserved", "chat.Components.Dispose"));
                Assert.AreEqual(observation.Root.InvocationId, row["InvocationId"]);
                Assert.IsNull(row["FaultCaught"]); Assert.IsNotNull(row["ErrorIsComException"]);
                StringAssert.Contains((string)row["Error"], failure.Message);
                StringAssert.Contains((string)row["Error"], "Observed callback stack:");
                Assert.IsFalse(observation.Events.Any(item => Is(item, "Fault", "chat.Components.Dispose")));
                int count = observation.Events.Count;
                try { throw new IOException("After observer retirement"); } catch (IOException) { }
                Assert.AreEqual(count, observation.Events.Count);
                Assert.IsNull(Get(window, "shutdownDiagnostic"));
            }
        }

        [STATestMethod]
        public void CaughtExceptionOnAnotherThreadIsOutsideTheAdmittedShutdownObservation()
        {
            var observation = new Observation();
            using (var window = new ChatWindow())
            {
                var original = (IContainer)Get(window, "components");
                var container = new CallbackContainer(original); Set(window, "components", container);
                container.OnDispose = () => {
                    var worker = new Thread(() => { try { throw new IOException("Another thread"); } catch (IOException) { } });
                    worker.Start(); Assert.IsTrue(worker.Join(5000));
                };
                window.DisposeForHostShutdown(observation.Root);
                Assert.IsTrue(window.IsDisposed); Assert.IsFalse(observation.Root.DiagnosticFailed);
                Assert.IsFalse(observation.Events.Any(row => (string)row["Phase"] == "ExceptionObserved"));
            }
        }

        [STATestMethod]
        public void FirstChancePublicationFailureIsBoundedAndDoesNotChangeOriginalCleanup()
        {
            var observation = new Observation(); int rejected = 0;
            observation.BeforePublish = row => {
                if ((string)row["Phase"] == "ExceptionObserved") { rejected++; throw new IOException("Observer publication failure"); }
            };
            using (var window = new ChatWindow())
            {
                var original = (IContainer)Get(window, "components");
                var container = new CallbackContainer(original); Set(window, "components", container);
                container.OnDispose = () => { try { throw new IOException("Original caught callback"); } catch (IOException) { } };
                window.DisposeForHostShutdown(observation.Root);
                Assert.AreEqual(1, rejected); Assert.AreEqual(1, container.Calls);
                Assert.IsTrue(window.IsDisposed); Assert.IsTrue(observation.Root.DiagnosticFailed);
                Assert.IsNull(Get(window, "shutdownDiagnostic"));
            }
        }

        [STATestMethod]
        public void NestedShutdownUsesOneFirstChanceObserverAndRestoresTheOuterInvocation()
        {
            var observation = new Observation();
            using (var window = new ChatWindow())
            {
                var original = (IContainer)Get(window, "components");
                var container = new CallbackContainer(original); Set(window, "components", container);
                container.OnDispose = () => {
                    window.DisposeForHostShutdown(observation.Root.Child("Dispose"));
                    try { throw new IOException("After nested shutdown returns"); } catch (IOException) { }
                };
                window.DisposeForHostShutdown(observation.Root);
                Assert.IsTrue(window.IsDisposed); Assert.IsFalse(observation.Root.DiagnosticFailed);
                var row = observation.Events.Single(item => (string)item["Phase"] == "ExceptionObserved");
                Assert.AreEqual(observation.Root.InvocationId, row["InvocationId"]);
                Assert.AreEqual("chat.Components.Dispose", row["Stage"]);
                Assert.IsNull(Get(window, "shutdownDiagnostic")); Assert.AreEqual(2, container.Calls);
            }
        }

        [STATestMethod]
        public void ActualReceiptPublicationDoesNotRecordItsExpectedMissingPathExceptions()
        {
            using (var directory = new ShutdownDiagnosticTestData.DirectoryScope())
            using (var window = new ChatWindow())
            {
                var observation = new Observation();
                AddInShutdownDiagnostic.Publish(directory.Root, 1, observation.Events[0], false);
                observation.BeforePublish = row => {
                    bool terminal = row["Stage"] == null && ((string)row["Phase"] == "Returned" || (string)row["Phase"] == "Fault");
                    AddInShutdownDiagnostic.Publish(directory.Root, (int)row["Sequence"], row, terminal);
                };
                window.DisposeForHostShutdown(observation.Root);
                Assert.IsTrue(window.IsDisposed); Assert.IsFalse(observation.Root.DiagnosticFailed);
                Assert.IsFalse(observation.Events.Any(row => (string)row["Phase"] == "ExceptionObserved"));
                Assert.AreEqual(observation.Events.Count, Directory.GetFiles(directory.Root, "*.json").Length);
                Assert.AreEqual(1, Directory.GetFiles(directory.Root, "*.ready").Length);
                Assert.AreEqual(0, Directory.GetFiles(directory.Root, "*.tmp").Length);
                Assert.IsNull(Get(window, "shutdownDiagnostic"));
            }
        }

    }
}
