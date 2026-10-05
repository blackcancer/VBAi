using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Web.Script.Serialization;
using System.Windows.Forms;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VBAi.Tests.Infrastructure;

namespace VBAi.Tests.Unit
{
    /// <summary>Mirrors the real shutdown call sites using managed resources and the pure publication seam.</summary>
    [TestClass, TestCategory("Unit"), DoNotParallelize]
    public sealed class AddInShutdownObservationTests
    {
        private const BindingFlags Private = BindingFlags.Instance | BindingFlags.NonPublic;
        private const string Nonce = "e0dc627ec00147c5b2304c503c144afa";
        private static readonly JavaScriptSerializer Json = new JavaScriptSerializer();

        private sealed class ObservationScope : IDisposable
        {
            private readonly Func<string, AddInShutdownDiagnostic> priorFactory = AddIn.CreateShutdownDiagnostic;
            private readonly Action priorStop = AddIn.StopUpdateCheck;
            private readonly Action<string> priorLog = AddIn.WriteLog;
            internal readonly List<Dictionary<string, object>> Events = new List<Dictionary<string, object>>();
            internal readonly List<AddInShutdownDiagnostic> Roots = new List<AddInShutdownDiagnostic>();
            internal readonly List<string> Logs = new List<string>();
            internal Action<Dictionary<string, object>> BeforePublish;
            internal int Stops;

            internal ObservationScope()
            {
                AddIn.WriteLog = Logs.Add;
                AddIn.StopUpdateCheck = () => Stops++;
                AddIn.CreateShutdownDiagnostic = Begin;
            }

            private AddInShutdownDiagnostic Begin(string entry)
            {
                var identity = new AddInShutdownDiagnostic.Identity
                {
                    ProcessId = 71, ProcessStartedUtc = "2026-10-05T17:27:03.6593947Z",
                    HostImagePath = @"C:\Qualification\EXCEL.EXE", ProductPath = @"C:\Qualification\VBAi.dll",
                    ProductMvid = "f855d463-0d5e-48f3-8c45-1a35db5a1025", ProductSha256 = new string('A', 64), ThreadId = 73
                };
                var diagnostic = AddInShutdownDiagnostic.Begin(Json.Serialize(new { Version = 1, Nonce, Identity = identity }),
                    Nonce, entry, () => identity,
                    () => new AddInShutdownDiagnostic.ThreadIdentity { ManagedThreadId = 1, NativeThreadId = 73, Apartment = "STA" },
                    (sequence, data, terminal) =>
                    {
                        var row = (Dictionary<string, object>)Json.DeserializeObject(Json.Serialize(data));
                        Assert.AreEqual(sequence, Convert.ToInt32(row["Sequence"]));
                        Assert.AreEqual(terminal, row["Stage"] == null && (string)row["Phase"] != "Entry");
                        BeforePublish?.Invoke(row);
                        Events.Add(row);
                    });
                Roots.Add(diagnostic);
                return diagnostic;
            }

            public void Dispose()
            {
                AddIn.CreateShutdownDiagnostic = priorFactory;
                AddIn.StopUpdateCheck = priorStop;
                AddIn.WriteLog = priorLog;
            }
        }

        /// <summary>Public so the actual dynamic call site can bind without a COM server.</summary>
        public sealed class ObservedNativeWindow
        {
            public int Calls;
            public Exception Error;
            public Action OnClose;
            public void Close() { Calls++; OnClose?.Invoke(); if (Error != null) throw Error; }
        }

        private static object Get(AddIn instance, string field) => typeof(AddIn).GetField(field, Private).GetValue(instance);
        private static void Set(AddIn instance, string field, object value) => typeof(AddIn).GetField(field, Private).SetValue(instance, value);
        private static void Shutdown(AddIn instance, bool begin)
        { object[] custom = null; if (begin) instance.OnBeginShutdown(ref custom); else instance.OnDisconnection(0, ref custom); }
        private static bool Is(Dictionary<string, object> row, string phase, string stage)
            => (string)row["Phase"] == phase && (string)row["Stage"] == stage;
        private static Dictionary<string, object>[] For(ObservationScope scope, string entry)
            => scope.Events.Where(row => (string)row["EntryPoint"] == entry).ToArray();
        private static void Before(ObservationScope scope, string firstPhase, string firstStage, string lastPhase, string lastStage)
        {
            int first = scope.Events.FindIndex(row => Is(row, firstPhase, firstStage));
            int last = scope.Events.FindIndex(row => Is(row, lastPhase, lastStage));
            Assert.IsTrue(first >= 0 && last > first, firstStage + " must precede " + lastStage);
        }

        [STATestMethod]
        [DataRow(false)] [DataRow(true)]
        public void NullResourcesHaveNoInventedCallsAndEachRealAssignmentFollowsItsWrite(bool begin)
        {
            using (var scope = new ObservationScope())
            {
                var instance = new AddIn();
                scope.BeforePublish = row => { if ((string)row["Phase"] == "ManagedReferenceCleared") Assert.IsNull(Get(instance, (string)row["Stage"])); };
                Shutdown(instance, begin);
                Assert.AreEqual(1, scope.Stops);
                var child = For(scope, "Dispose");
                CollectionAssert.AreEqual(new[] { "VbeNativeTheme.Disconnect", "StopUpdateCheck" },
                    child.Where(row => (string)row["Phase"] == "Entry" && row["Stage"] != null).Select(row => (string)row["Stage"]).ToArray());
                CollectionAssert.AreEqual(new[] { "editorNavigation", "editorWorkspace", "modernEditor", "testExplorerWindow", "nativeTestWindow", "nativeTestControl",
                    "testExplorerService", "crashReporter", "menu", "chat", "nativeChatWindow", "nativeChatControl", "addIn", "server", "dispatcher", "vbe" },
                    child.Where(row => (string)row["Phase"] == "ManagedReferenceCleared").Select(row => (string)row["Stage"]).ToArray());
                var parent = For(scope, begin ? "OnBeginShutdown" : "OnDisconnection");
                Assert.AreEqual(2, parent.Length);
                Assert.IsTrue(Is(parent[0], "Entry", null)); Assert.IsTrue(Is(parent[1], "Returned", null));
                Assert.AreEqual(parent[0]["InvocationId"], child[0]["ParentInvocationId"]);
                Assert.AreNotEqual(parent[0]["InvocationId"], child[0]["InvocationId"]);
                Assert.IsFalse(scope.Roots.Single().DiagnosticFailed);
            }
        }

        [STATestMethod]
        [DataRow(false)] [DataRow(true)]
        public void RepeatedCallbacksHaveDistinctParentsAndCloseOnlyTheStillPresentWindow(bool firstBegin)
        {
            using (var scope = new ObservationScope())
            using (var dispatcher = new Control())
            {
                var instance = new AddIn(); var native = new ObservedNativeWindow();
                Set(instance, "nativeChatWindow", native); Set(instance, "dispatcher", dispatcher);
                scope.BeforePublish = row =>
                {
                    if (Is(row, "Returned", "nativeChatWindow.Close")) { Assert.AreEqual(1, native.Calls); Assert.AreSame(native, Get(instance, "nativeChatWindow")); }
                    if (Is(row, "ManagedReferenceCleared", "nativeChatWindow")) Assert.IsNull(Get(instance, "nativeChatWindow"));
                };
                Shutdown(instance, firstBegin); Shutdown(instance, !firstBegin);
                Assert.AreEqual(1, native.Calls); Assert.AreEqual(2, scope.Stops); Assert.IsTrue(dispatcher.IsDisposed);
                Assert.AreEqual(2, scope.Roots.Count); Assert.AreNotEqual(scope.Roots[0].InvocationId, scope.Roots[1].InvocationId);
                var children = For(scope, "Dispose").Where(row => Is(row, "Entry", null)).ToArray();
                CollectionAssert.AreEqual(scope.Roots.Select(root => root.InvocationId).ToArray(), children.Select(row => (string)row["ParentInvocationId"]).ToArray());
                Assert.AreEqual(2, children.Select(row => row["InvocationId"]).Distinct().Count());
                Before(scope, "Returned", "nativeChatWindow.Close", "ManagedReferenceCleared", "nativeChatWindow");
                Before(scope, "Returned", "dispatcher.Dispose", "ManagedReferenceCleared", "dispatcher");
                Assert.IsTrue(scope.Roots.All(root => !root.DiagnosticFailed));
            }
        }

        [STATestMethod]
        public void DirectPrivateDisposeKeepsItsUniqueSignatureAndOwnRootInvocation()
        {
            using (var scope = new ObservationScope())
            {
                var method = typeof(AddIn).GetMethod("Dispose", Private);
                Assert.AreEqual(0, method.GetParameters().Length);
                method.Invoke(new AddIn(), null);
                Assert.AreEqual(1, scope.Roots.Count);
                Assert.IsTrue(scope.Events.All(row => row["ParentInvocationId"] == null && (string)row["EntryPoint"] == "Dispose"));
                Assert.IsTrue(Is(scope.Events.Last(), "Returned", null));
            }
        }

        [STATestMethod]
        [DataRow(false, 0)] [DataRow(false, 1)] [DataRow(false, 2)]
        [DataRow(true, 0)] [DataRow(true, 1)] [DataRow(true, 2)]
        public void NativeClosePreservesTheTestComOnlyAndChatUniversalCatchPolicies(bool testWindow, int failure)
        {
            using (var scope = new ObservationScope())
            using (var dispatcher = new Control())
            {
                var instance = new AddIn();
                Exception primary = failure == 0 ? null : failure == 1 ? (Exception)new COMException("close rejected", unchecked((int)0x80010001)) : new InvalidOperationException("close failed");
                var native = new ObservedNativeWindow { Error = primary };
                string field = testWindow ? "nativeTestWindow" : "nativeChatWindow", stage = field + ".Close";
                Set(instance, field, native); Set(instance, "dispatcher", dispatcher);
                bool propagates = testWindow && failure == 2;
                if (propagates) Assert.AreSame(primary, Assert.ThrowsException<InvalidOperationException>(() => Shutdown(instance, true)));
                else Shutdown(instance, true);
                Assert.AreEqual(1, native.Calls);
                Assert.AreEqual(!propagates, dispatcher.IsDisposed);
                Assert.AreEqual(propagates ? 0 : 1, scope.Stops);
                Assert.AreEqual(failure == 0 ? 1 : 0, scope.Events.Count(row => Is(row, "Returned", stage)));
                Assert.AreEqual(failure == 0 ? 0 : 1, scope.Events.Count(row => Is(row, "Fault", stage)));
                Assert.AreEqual(propagates ? 2 : 0, scope.Events.Count(row => Is(row, "Fault", null)));
                if (propagates) { Assert.AreSame(native, Get(instance, field)); Assert.IsFalse(scope.Events.Any(row => Is(row, "ManagedReferenceCleared", field))); }
                else { Assert.IsNull(Get(instance, field)); Before(scope, failure == 0 ? "Returned" : "Fault", stage, "ManagedReferenceCleared", field); }
                Assert.IsFalse(scope.Roots.Single().DiagnosticFailed);
            }
        }

        [STATestMethod]
        [DataRow(false)] [DataRow(true)]
        public void StopUpdateFailureKeepsTheSameExceptionAndPreventsLaterChatServerAndDispatcherCleanup(bool begin)
        {
            using (var host = new HostUiScope())
            using (var scope = new ObservationScope())
            {
                var instance = host.Connected();
                var chat = Get(instance, "chat"); var server = Get(instance, "server"); var dispatcher = (Control)Get(instance, "dispatcher");
                Assert.IsNotNull(chat); Assert.IsNotNull(server);
                var primary = new IOException("stop update failed"); int calls = 0;
                AddIn.StopUpdateCheck = () => { calls++; throw primary; };
                try
                {
                    Assert.AreSame(primary, Assert.ThrowsException<IOException>(() => Shutdown(instance, begin)));
                    Assert.AreEqual(1, calls); Assert.AreSame(chat, Get(instance, "chat")); Assert.AreSame(server, Get(instance, "server"));
                    Assert.AreSame(dispatcher, Get(instance, "dispatcher")); Assert.IsFalse(dispatcher.IsDisposed);
                    Assert.IsFalse(scope.Events.Any(row => Is(row, "Entry", "chat.Dispose") || Is(row, "Entry", "server.Dispose") || Is(row, "Entry", "dispatcher.Dispose")));
                    Assert.AreEqual(1, scope.Events.Count(row => Is(row, "Fault", "StopUpdateCheck")));
                    Assert.AreEqual(2, scope.Events.Count(row => Is(row, "Fault", null)));
                    Assert.IsFalse(scope.Events.Any(row => Is(row, "Returned", null)));
                }
                finally { AddIn.StopUpdateCheck = () => { }; AddIn.CreateShutdownDiagnostic = entry => null; host.Close(instance); }
            }
        }

        [STATestMethod]
        [DataRow("RootEntry", false)] [DataRow("RootEntry", true)]
        [DataRow("ChildEntry", false)] [DataRow("ChildEntry", true)]
        [DataRow("StageEntry", false)] [DataRow("StageEntry", true)]
        [DataRow("StageReturn", false)] [DataRow("StageReturn", true)]
        [DataRow("Clear", false)] [DataRow("Clear", true)]
        [DataRow("Terminal", false)] [DataRow("Terminal", true)]
        public void PublicationFailureCannotReplaceTheOriginalExceptionOrSkipAnOriginalCall(string boundary, bool nativeFailure)
        {
            using (var scope = new ObservationScope())
            using (var dispatcher = new Control())
            {
                var instance = new AddIn(); var primary = new IOException("original dispatcher disposal");
                var publication = new IOException("diagnostic publication");
                var native = new ObservedNativeWindow(); Set(instance, "nativeChatWindow", native); Set(instance, "dispatcher", dispatcher);
                EventHandler disposed = (sender, args) => { throw primary; };
                if (nativeFailure) dispatcher.Disposed += disposed;
                int attempts = 0, failedAt = 0;
                scope.BeforePublish = row =>
                {
                    attempts++;
                    bool selected = boundary == "RootEntry" ? Is(row, "Entry", null) && row["ParentInvocationId"] == null :
                        boundary == "ChildEntry" ? Is(row, "Entry", null) && row["ParentInvocationId"] != null :
                        boundary == "StageEntry" ? Is(row, "Entry", "nativeChatWindow.Close") :
                        boundary == "StageReturn" ? Is(row, "Returned", "nativeChatWindow.Close") :
                        boundary == "Clear" ? Is(row, "ManagedReferenceCleared", "nativeChatWindow") : row["Stage"] == null && (string)row["Phase"] != "Entry";
                    if (selected) { failedAt = attempts; throw publication; }
                };
                try
                {
                    if (nativeFailure) Assert.AreSame(primary, Assert.ThrowsException<IOException>(() => Shutdown(instance, true)));
                    else Shutdown(instance, true);
                }
                finally { dispatcher.Disposed -= disposed; }
                Assert.AreEqual(1, native.Calls); Assert.AreEqual(1, scope.Stops);
                Assert.IsNull(Get(instance, "nativeChatWindow"));
                Assert.IsTrue(failedAt > 0); Assert.AreEqual(failedAt, attempts, "Publication must stop after its first fault.");
                Assert.IsTrue(scope.Roots.Single().DiagnosticFailed); Assert.AreSame(publication, scope.Roots.Single().PublicationError);
                Assert.AreEqual(0, scope.Events.Count(row => row["Stage"] == null && (string)row["Phase"] != "Entry"));
            }
        }

        [STATestMethod]
        [DataRow(false)] [DataRow(true)]
        public void AdmissionSeamFailureDoesNotPreventCleanupOrReplaceItsException(bool nativeFailure)
        {
            using (var scope = new ObservationScope())
            using (var dispatcher = new Control())
            {
                var instance = new AddIn(); var native = new ObservedNativeWindow(); var primary = new IOException("original stop");
                Set(instance, "nativeChatWindow", native); Set(instance, "dispatcher", dispatcher);
                AddIn.CreateShutdownDiagnostic = entry => { throw new IOException("admission"); };
                if (nativeFailure) AddIn.StopUpdateCheck = () => { throw primary; };
                if (nativeFailure) Assert.AreSame(primary, Assert.ThrowsException<IOException>(() => Shutdown(instance, false)));
                else Shutdown(instance, false);
                Assert.AreEqual(nativeFailure ? 0 : 1, native.Calls); Assert.AreEqual(!nativeFailure, dispatcher.IsDisposed);
                Assert.AreEqual(0, scope.Events.Count);
            }
        }

        [STATestMethod]
        [DataRow(false)] [DataRow(true)]
        public void NullAdmissionKeepsBothOriginalCallbackPaths(bool begin)
        {
            using (var scope = new ObservationScope())
            using (var dispatcher = new Control())
            {
                AddIn.CreateShutdownDiagnostic = entry => null;
                var instance = new AddIn(); var native = new ObservedNativeWindow();
                Set(instance, "nativeChatWindow", native); Set(instance, "dispatcher", dispatcher);
                Shutdown(instance, begin);
                Assert.AreEqual(1, native.Calls); Assert.AreEqual(1, scope.Stops); Assert.IsTrue(dispatcher.IsDisposed);
                Assert.AreEqual(0, scope.Events.Count); Assert.AreEqual(0, scope.Roots.Count);
            }
        }

        [STATestMethod]
        public void ReentrantCallbackKeepsSeparateInvocationIdentitiesAndExistingCloseBehavior()
        {
            using (var scope = new ObservationScope())
            {
                var instance = new AddIn(); var native = new ObservedNativeWindow();
                Set(instance, "nativeChatWindow", native);
                native.OnClose = () => { native.OnClose = null; Shutdown(instance, false); };
                Shutdown(instance, true);
                Assert.AreEqual(2, native.Calls, "Observation must not add a cleanup guard or suppress the existing reentrant call.");
                Assert.AreEqual(2, scope.Roots.Count); Assert.AreEqual(2, scope.Stops);
                var starts = scope.Events.Where(row => Is(row, "Entry", null)).ToArray();
                Assert.AreEqual(4, starts.Select(row => row["InvocationId"]).Distinct().Count());
                foreach (var root in scope.Roots)
                    Assert.AreEqual(1, starts.Count(row => (string)row["ParentInvocationId"] == root.InvocationId));
                Assert.AreEqual(4, scope.Events.Count(row => Is(row, "Returned", null)));
                Assert.IsTrue(scope.Roots.All(root => !root.DiagnosticFailed));
            }
        }

        [STATestMethod]
        [DataRow(false)] [DataRow(true)]
        public void ThemeFaultRemainsCaughtAndCannotHideALaterOriginalFailure(bool failStop)
        {
            using (var scope = new ObservationScope())
            using (var dispatcher = new Control())
            {
                var hook = typeof(VbeNativeTheme).GetField("windowEventHook", BindingFlags.Static | BindingFlags.NonPublic);
                object priorHook = hook.GetValue(null); var priorRemove = VbeNativeTheme.RemoveWindowHook;
                Assert.AreEqual(IntPtr.Zero, (IntPtr)priorHook, "The managed mirror must not replace an installed hook.");
                var themeError = new IOException("theme remove failed"); var stopError = new IOException("original stop failed");
                var instance = new AddIn(); Set(instance, "dispatcher", dispatcher); int removes = 0;
                try
                {
                    hook.SetValue(null, new IntPtr(123));
                    VbeNativeTheme.RemoveWindowHook = handle => { Assert.AreEqual(new IntPtr(123), handle); removes++; throw themeError; };
                    if (failStop) AddIn.StopUpdateCheck = () => { scope.Stops++; throw stopError; };
                    if (failStop) Assert.AreSame(stopError, Assert.ThrowsException<IOException>(() => Shutdown(instance, true)));
                    else Shutdown(instance, true);
                    Assert.AreEqual(1, removes); Assert.AreEqual(1, scope.Stops); Assert.AreEqual(!failStop, dispatcher.IsDisposed);
                    Assert.AreEqual(1, scope.Events.Count(row => Is(row, "Fault", "VbeNativeTheme.Disconnect")));
                    Assert.IsFalse(scope.Events.Any(row => Is(row, "Returned", "VbeNativeTheme.Disconnect")));
                    Assert.IsTrue(scope.Logs.Any(log => log.Contains(themeError.Message)));
                    Assert.AreEqual(failStop ? 2 : 0, scope.Events.Count(row => Is(row, "Fault", null)));
                    Assert.IsFalse(scope.Roots.Single().DiagnosticFailed);
                }
                finally { hook.SetValue(null, priorHook); VbeNativeTheme.RemoveWindowHook = priorRemove; }
            }
        }

        [STATestMethod]
        [DataRow(false)] [DataRow(true)]
        public void ToolbarCleanupRecordsItsExistingCaughtFaultAndStillDisposes(bool rejectDelete)
        {
            using (var scope = new ObservationScope())
            using (var dispatcher = new Control())
            {
                var host = new ToolbarCustomizationTests.Host(); var bar = host.CommandBars.Add("Standard", 1, false, true);
                var button = bar.Controls.Add(1, 42, Type.Missing, 1, true);
                button.Tag = "VBAi.ToolbarCommand." + new string('a', 32); button.FailDelete = rejectDelete;
                var instance = new AddIn(); Set(instance, "vbe", host); Set(instance, "dispatcher", dispatcher);
                Shutdown(instance, true);
                Assert.IsTrue(dispatcher.IsDisposed); Assert.AreEqual(rejectDelete ? 1 : 0, bar.Controls.Count);
                Assert.AreEqual(1, scope.Events.Count(row => Is(row, rejectDelete ? "Fault" : "Returned", "CleanupTemporaryToolbarCommands")));
                Assert.AreEqual(0, scope.Events.Count(row => Is(row, rejectDelete ? "Returned" : "Fault", "CleanupTemporaryToolbarCommands")));
                Assert.IsFalse(scope.Roots.Single().DiagnosticFailed);
            }
        }

        [STATestMethod]
        [DataRow(false)] [DataRow(true)]
        public void ChatDetachmentIsObservedOnlyForTheExistingDockedOwnerCondition(bool docked)
        {
            using (var host = new HostUiScope())
            using (var scope = new ObservationScope())
            {
                var instance = host.Connected();
                try
                {
                    if (!docked) typeof(AddIn).GetMethod("ToggleDock", Private).Invoke(instance, null);
                    Assert.AreEqual(docked, (bool)Get(instance, "docked"));
                    Shutdown(instance, true);
                    Assert.AreEqual(docked ? 1 : 0, scope.Events.Count(row => Is(row, "Entry", "nativeChatControl.Detach")));
                    Assert.AreEqual(docked ? 1 : 0, scope.Events.Count(row => Is(row, "Returned", "nativeChatControl.Detach")));
                    Before(scope, "Returned", "chat.Dispose", "ManagedReferenceCleared", "chat");
                    Before(scope, "Returned", "server.Dispose", "ManagedReferenceCleared", "server");
                    Assert.IsFalse(scope.Roots.Single().DiagnosticFailed);
                }
                finally { AddIn.CreateShutdownDiagnostic = entry => null; host.Close(instance); }
            }
        }

        [STATestMethod]
        [DataRow(false)] [DataRow(true)]
        public void TestWindowDetachmentIsObservedOnlyForAnAttachedLiveWindow(bool attached)
        {
            using (var scope = new ObservationScope())
            using (var container = new ChatToolWindow())
            using (var window = new TestExplorerWindow())
            {
                var instance = new AddIn();
                if (attached) { Assert.AreNotEqual(IntPtr.Zero, container.Handle); container.Attach(window); }
                Set(instance, "testExplorerWindow", window); Set(instance, "nativeTestControl", container);
                Shutdown(instance, true);
                Assert.AreEqual(attached ? 1 : 0, scope.Events.Count(row => Is(row, "Returned", "nativeTestControl.Detach")));
                Before(scope, "Returned", "testExplorerWindow.Dispose", "ManagedReferenceCleared", "testExplorerWindow");
                Assert.IsTrue(window.IsDisposed); Assert.IsFalse(scope.Roots.Single().DiagnosticFailed);
            }
        }
    }
}
