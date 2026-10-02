using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Web.Script.Serialization;
using System.Windows.Automation;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Integration
{
    /// <summary>Exercises only the installed Word chat scope picker and its GitHub modal entry.</summary>
    [TestClass, TestCategory("Office"), TestCategory("NativeWordChatGitUi"), DoNotParallelize]
    public sealed class WordChatGitWindowTests
    {
        private static readonly List<Context> Retained = new List<Context>();

        [TestMethod]
        public void SavedWordChatScopeOpensOnlyItsOwnedGitModal()
        {
            if (Environment.GetEnvironmentVariable("VBAi_RUN_WORD_CHAT_GIT_TESTS") != "1" ||
                Environment.GetEnvironmentVariable("VBAi_RUN_WORD_GIT_TESTS") != "1" ||
                Environment.GetEnvironmentVariable("VBAi_RUN_OFFICE_TESTS") != "1")
                Assert.Inconclusive("Owned Word and chat-to-Git UI opt-ins are required.");
            string root = ExcelOwnedBootstrapPlan.RequireLocalAbsolutePath(Environment.GetEnvironmentVariable("VBAi_OFFICE_RESULTS"));
            Guid expected = Guid.Parse(Environment.GetEnvironmentVariable("VBAi_TEST_EMBEDDED_GIT_MVID"));
            string hash = Environment.GetEnvironmentVariable("VBAi_TEST_EMBEDDED_GIT_SHA256");
            Assert.AreEqual(expected, typeof(VbeSession).Module.ModuleVersionId);
            Assert.AreEqual(hash, Sha(typeof(VbeSession).Assembly.Location), true);
            var context = new Context(Path.Combine(root, "word-chat-git-" + Guid.NewGuid().ToString("N")));
            context.Record(new { Phase = "Preflight", ExpectedMvid = expected.ToString("D"), ExpectedSha256 = hash,
                Scope = "Actual owned Word chat scope selection and Git modal opening/close only; no provider prompt, Git connect, remote write, macro or import." });
            var owner = new Thread(() => Owner(context, expected, hash)) { IsBackground = true };
            owner.SetApartmentState(ApartmentState.STA); owner.Start();
            try
            {
                if (!context.Ready.Wait(TimeSpan.FromSeconds(90))) throw new TimeoutException("Owned Word preparation did not reach a terminal state.");
                if (context.OwnerError != null) ExceptionDispatchInfo.Capture(context.OwnerError).Throw();
                var ui = new Thread(() => Ui(context)) { IsBackground = true };
                ui.SetApartmentState(ApartmentState.MTA); ui.Start();
                if (!context.UiDone.Wait(TimeSpan.FromSeconds(120))) throw new TimeoutException("Chat/Git UI observation is pending or uncertain.");
                if (!context.OwnerDone.Wait(TimeSpan.FromSeconds(30))) throw new TimeoutException("Owned Word cleanup has no terminal evidence.");
                var errors = new[] { context.UiError, context.OwnerError }.Where(error => error != null).ToArray();
                if (errors.Length > 1) throw new AggregateException("Word chat UI and owner failures are preserved.", errors);
                if (errors.Length == 1) ExceptionDispatchInfo.Capture(errors[0]).Throw();
                context.Record(new { Phase = "PASS", ProcessId = context.Fixture.ProcessId,
                    Scope = "WordChatScopeGitModalKnownCloseNormalExit" });
            }
            catch (Exception primary)
            {
                context.Stop = true;
                context.ScopeObserved.Set(); context.ScopeAuthorized.Set();
                if (!context.ActionIssued && !context.Retain && context.Ready.IsSet)
                    context.OwnerDone.Wait(TimeSpan.FromSeconds(30));
                if (!context.OwnerDone.IsSet)
                {
                    context.Retain = true;
                    if (context.Fixture != null) context.Fixture.NativeExecutionUnsettled = true;
                    lock (Retained) if (!Retained.Contains(context)) Retained.Add(context);
                }
                Exception preserved = EmbeddedGitUiProtocol.PreserveFailures(primary, context.UiError, context.OwnerError);
                context.Record(new { Phase = "FailedOrRetained", context.Retain, context.ActionIssued,
                    context.ModalObserved, context.ModalClosed, OwnerTerminal = context.OwnerDone.IsSet,
                    UiTerminal = context.UiDone.IsSet, Error = preserved.ToString(), ReplayAttempts = 0 });
                if (!ReferenceEquals(primary, preserved)) ExceptionDispatchInfo.Capture(preserved).Throw();
                throw;
            }
        }

        private static void Owner(Context context, Guid expected, string hash)
        {
            try
            {
                context.Fixture = OfficeVbeFixture.Start("Word");
                var status = context.Fixture.Data("status");
                Assert.AreEqual(expected.ToString("D"), status["AssemblyModuleVersionId"]);
                Assert.AreEqual(hash, Sha(Convert.ToString(status["AssemblyPath"])), true);
                context.Scope = context.Fixture.PrepareWordEmbeddedGit(context.Nonce, context.Record);
                context.Label = context.Fixture.RequireWordChatGitScope(context.Scope);
                context.DocumentHash = Sha(context.Scope.Path);
                Assert.IsFalse(Directory.Exists(context.Scope.Cache), "A new Word chat scope must have no prior Git binding.");
                context.Record(new { Phase = "Prepared", context.Fixture.ProcessId, context.Scope.Path,
                    context.Scope.ThreadId, context.Label, context.DocumentHash });
                context.Ready.Set();
                if (!context.ScopeObserved.Wait(TimeSpan.FromSeconds(45))) throw new TimeoutException("Chat scope selection was not observed.");
                if (context.Stop || context.UiError != null) throw new InvalidOperationException("Chat UI stopped before Git action.");
                Assert.AreEqual(context.Label, context.Fixture.RequireWordChatGitScope(context.Scope));
                context.Record(new { Phase = "ChatScopeAuthorized", CanonicalPath = context.Scope.Path, context.Label });
                context.ScopeAuthorized.Set();
                if (!context.UiDone.Wait(TimeSpan.FromSeconds(100))) throw new TimeoutException("Chat Git modal did not reach known closure.");
                if (context.UiError != null || !context.ModalClosed) throw new InvalidOperationException("Chat Git modal did not complete exactly.");
                context.Fixture.VerifyWordEmbeddedSource(context.Scope);
                Assert.AreEqual(context.DocumentHash, Sha(context.Scope.Path));
                Assert.IsTrue(Directory.Exists(context.Scope.Cache), "The owned Git modal must initialize only its selected document cache.");
                Assert.IsFalse(File.Exists(Path.Combine(context.Scope.Cache, "binding.json")), "Chat opening must not link a remote.");
                Assert.AreEqual(0, Directory.EnumerateDirectories(context.Scope.Cache, "*.git", SearchOption.TopDirectoryOnly).Count(),
                    "Opening the chat Git modal must not initialize a repository.");
                context.Record(new { Phase = "WordChatGitSourcePreserved", context.DocumentHash, RemoteWrites = 0 });
            }
            catch (Exception error) { context.OwnerError = error; }
            finally
            {
                context.Ready.Set(); context.ScopeAuthorized.Set();
                if (context.ActionIssued && !context.ModalClosed) context.Retain = true;
                if (context.Fixture != null)
                {
                    if (context.Fixture.WordGitMustRetain) context.Retain = true;
                    if (context.Retain)
                    {
                        context.Fixture.NativeExecutionUnsettled = true;
                        lock (Retained) if (!Retained.Contains(context)) Retained.Add(context);
                    }
                    else
                    {
                        try
                        {
                            context.Record(new { Phase = "NormalCleanupIntent", ReplayAttempts = 0 });
                            context.Fixture.Dispose();
                            if (context.DocumentHash != null) Assert.AreEqual(context.DocumentHash, Sha(context.Scope.Path));
                            context.Record(new { Phase = "NormalCleanupReturned" });
                        }
                        catch (Exception cleanup) { context.OwnerError = context.OwnerError == null ? cleanup : new AggregateException(context.OwnerError, cleanup); }
                    }
                }
                context.OwnerDone.Set();
            }
        }

        private static void Ui(Context context)
        {
            var automation = new WordChatGitAutomation(context);
            try
            {
                context.Record(new { Phase = "MtaUiAArmed", Apartment = Thread.CurrentThread.GetApartmentState().ToString() });
                automation.FindAndSelectScope();
                context.ScopeObserved.Set();
                if (!context.ScopeAuthorized.Wait(TimeSpan.FromSeconds(30)) || context.Stop || context.OwnerError != null)
                    throw new InvalidOperationException("Canonical Word scope was not authorized before chat Git invocation.");
                automation.OpenChatOptionsAndFindGit();
                automation.RequireNoGitModal();
                context.InvocationThread = new Thread(() => automation.InvokeGitOnce()) { IsBackground = true };
                context.InvocationThread.SetApartmentState(ApartmentState.MTA); context.InvocationThread.Start();
                if (!context.GitIntent.Wait(TimeSpan.FromSeconds(10))) throw new TimeoutException("No durable chat Git invocation intent.");
                automation.ObserveGitModal();
                if (context.InvokerDone.IsSet && context.InvokerError != null) throw context.InvokerError;
                automation.CloseExactGitModal();
                if (!context.InvokerDone.Wait(TimeSpan.FromSeconds(15))) throw new TimeoutException("The chat Git invocation did not return after modal close.");
                if (context.InvokerError != null) throw context.InvokerError;
                context.Record(new { Phase = "ChatGitOpenCloseTerminal", context.ChatHandle, context.GitHandle,
                    context.Scope.Path, ReplayAttempts = 0 });
                context.ModalClosed = true;
                context.Fixture.NativeExecutionUnsettled = false;
            }
            catch (Exception error) { context.UiError = error; }
            finally { context.ScopeObserved.Set(); context.UiDone.Set(); }
        }

        private static string Sha(string path)
        {
            using (var file = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var hash = SHA256.Create()) return BitConverter.ToString(hash.ComputeHash(file)).Replace("-", "");
        }

        private sealed class Context
        {
            internal readonly string Root, Nonce = "WORD_CHAT_GIT_" + Guid.NewGuid().ToString("N");
            internal readonly ManualResetEventSlim Ready = new ManualResetEventSlim(), ScopeObserved = new ManualResetEventSlim(),
                ScopeAuthorized = new ManualResetEventSlim(), GitIntent = new ManualResetEventSlim(),
                InvokerDone = new ManualResetEventSlim(), UiDone = new ManualResetEventSlim(), OwnerDone = new ManualResetEventSlim();
            internal OfficeVbeFixture Fixture;
            internal ExcelVbeFixture.EmbeddedGitScope Scope;
            internal Thread InvocationThread;
            internal string Label, DocumentHash;
            internal IntPtr ChatHandle, GitHandle;
            internal Exception UiError, OwnerError, InvokerError;
            internal volatile bool Stop, Retain, ActionIssued, ModalObserved, ModalClosed;
            private readonly object sync = new object();
            private int sequence;
            internal Context(string root) { Root = root; Directory.CreateDirectory(root); }
            internal void Record(object value)
            {
                lock (sync)
                {
                    if (++sequence > 160) throw new InvalidOperationException("Word chat Git evidence bound exhausted.");
                    using (var file = new FileStream(Path.Combine(Root, sequence.ToString("D3") + ".json"), FileMode.CreateNew, FileAccess.Write, FileShare.Read))
                    using (var writer = new StreamWriter(file, new UTF8Encoding(false))) writer.Write(new JavaScriptSerializer().Serialize(value));
                }
            }
        }

        private sealed class WordChatGitAutomation
        {
            private delegate bool Visitor(IntPtr window, IntPtr data);
            [DllImport("user32.dll")] private static extern bool EnumWindows(Visitor visitor, IntPtr data);
            [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr window, out uint pid);
            [DllImport("user32.dll")] private static extern bool IsChild(IntPtr parent, IntPtr child);
            [DllImport("user32.dll")] private static extern bool IsWindow(IntPtr window);
            [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr window);
            [DllImport("user32.dll")] private static extern IntPtr GetWindow(IntPtr window, uint command);
            [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetClassName(IntPtr window, StringBuilder text, int capacity);
            [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetWindowText(IntPtr window, StringBuilder text, int capacity);
            private readonly Context context;
            private AutomationElement chat, scopePicker, gitItem, git;
            private SelectionPattern scopeSelection;
            private WindowPattern gitWindow;
            internal WordChatGitAutomation(Context context) { this.context = context; }

            private IntPtr[] OwnedTopWindows()
            {
                var found = new List<IntPtr>(); int visited = 0;
                Visitor visitor = (window, unused) => {
                    if (++visited > 4096) return false;
                    uint pid; GetWindowThreadProcessId(window, out pid);
                    if (pid == context.Fixture.ProcessId) found.Add(window);
                    return true;
                };
                if (!EnumWindows(visitor, IntPtr.Zero) || visited > 4096 || found.Count > 64)
                    throw new InvalidOperationException("Bounded Word top-window inventory failed.");
                return found.ToArray();
            }

            private static AutomationElement[] Descendants(AutomationElement root, string id)
                => root.FindAll(TreeScope.Descendants, new PropertyCondition(AutomationElement.AutomationIdProperty, id))
                    .Cast<AutomationElement>().ToArray();

            private void Guard(IntPtr window)
            {
                if (context.Stop) throw new InvalidOperationException("Coordinator stopped Word chat UI actions.");
                context.Fixture.RequireWordEmbeddedOwner(context.Scope);
                uint pid; uint tid = GetWindowThreadProcessId(window, out pid);
                EmbeddedGitUiProtocol.RequireOwner(context.Fixture.ProcessId, context.Scope.ThreadId,
                    window.ToInt64(), (int)pid, tid, window.ToInt64());
            }

            internal void FindAndSelectScope()
            {
                var matches = new List<AutomationElement>();
                foreach (IntPtr window in OwnedTopWindows())
                {
                    var root = AutomationElement.FromHandle(window);
                    if (root.Current.AutomationId == "ChatWindow") matches.Add(root);
                    matches.AddRange(Descendants(root, "ChatWindow"));
                }
                var unique = matches.GroupBy(item => item.Current.NativeWindowHandle).ToArray();
                if (unique.Length != 1 || unique[0].Key == 0) throw new InvalidOperationException("The exact owned Word chat window is absent or ambiguous.");
                chat = unique[0].First(); context.ChatHandle = new IntPtr(unique[0].Key);
                Guard(context.ChatHandle);
                if (!IsWindowVisible(context.ChatHandle) ||
                    !IsChild(context.Scope.VbeHandle, context.ChatHandle) && GetWindow(context.ChatHandle, 4) != context.Scope.VbeHandle)
                    throw new InvalidOperationException("The Word chat is not visible inside or owned by the exact VBE.");
                var pickers = Descendants(chat, "scopePicker");
                if (pickers.Length != 1 || pickers[0].Current.ControlType != ControlType.ComboBox ||
                    pickers[0].Current.ProcessId != context.Fixture.ProcessId)
                    throw new InvalidOperationException("The owned chat scope picker is absent or ambiguous.");
                scopePicker = pickers[0];
                scopeSelection = Pattern<SelectionPattern>(scopePicker, SelectionPattern.Pattern);
                if (!SelectedLabel(scopeSelection, context.Label))
                {
                    Pattern<ExpandCollapsePattern>(scopePicker, ExpandCollapsePattern.Pattern).Expand();
                    var entries = scopePicker.FindAll(TreeScope.Descendants,
                        new AndCondition(new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.ListItem),
                            new PropertyCondition(AutomationElement.NameProperty, context.Label))).Cast<AutomationElement>().ToArray();
                    if (entries.Length != 1 || entries[0].Current.ProcessId != context.Fixture.ProcessId)
                        throw new InvalidOperationException("The unique canonical Word chat scope has no selectable UI item.");
                    Pattern<SelectionItemPattern>(entries[0], SelectionItemPattern.Pattern).Select();
                    Pattern<ExpandCollapsePattern>(scopePicker, ExpandCollapsePattern.Pattern).Collapse();
                }
                var watch = Stopwatch.StartNew(); int stable = 0;
                while (watch.ElapsedMilliseconds < 15000 && stable < 2)
                {
                    Guard(context.ChatHandle);
                    stable = scopePicker.Current.IsEnabled && SelectedLabel(scopeSelection, context.Label) ? stable + 1 : 0;
                    if (stable < 2) Thread.Sleep(50);
                }
                if (stable != 2) throw new TimeoutException("The selected saved Word chat scope did not become idle and exact.");
                context.Record(new { Phase = "ChatScopeSelected", context.Label, CanonicalPath = context.Scope.Path,
                    ChatHandle = context.ChatHandle.ToInt64(), PickerHandle = scopePicker.Current.NativeWindowHandle });
            }

            private void RequireSelectedScope()
            {
                Guard(context.ChatHandle);
                if (scopePicker == null || scopeSelection == null || !scopePicker.Current.IsEnabled ||
                    !SelectedLabel(scopeSelection, context.Label))
                    throw new InvalidOperationException("The exact saved Word chat scope changed before Git invocation.");
            }

            private static bool SelectedLabel(SelectionPattern selection, string label)
            {
                var items = selection.Current.GetSelection();
                return items.Length == 1 && string.Equals(items[0].Current.Name, label, StringComparison.Ordinal);
            }

            internal void OpenChatOptionsAndFindGit()
            {
                RequireSelectedScope();
                var buttons = Descendants(chat, "options");
                if (buttons.Length != 1 || buttons[0].Current.ControlType != ControlType.Button ||
                    !buttons[0].Current.IsEnabled || buttons[0].Current.IsOffscreen)
                    throw new InvalidOperationException("The exact Word chat Options button is unavailable.");
                context.Record(new { Phase = "ChatOptionsIntent", ChatHandle = context.ChatHandle.ToInt64() });
                Pattern<InvokePattern>(buttons[0], InvokePattern.Pattern).Invoke();
                var watch = Stopwatch.StartNew();
                while (watch.ElapsedMilliseconds < 5000 && gitItem == null)
                {
                    var matches = new List<AutomationElement>();
                    foreach (IntPtr window in OwnedTopWindows())
                    {
                        uint pid; uint tid = GetWindowThreadProcessId(window, out pid);
                        if (tid != context.Scope.ThreadId || !IsWindowVisible(window)) continue;
                        var cls = new StringBuilder(128); GetClassName(window, cls, cls.Capacity);
                        if (!cls.ToString().StartsWith("WindowsForms", StringComparison.Ordinal)) continue;
                        matches.AddRange(Descendants(AutomationElement.FromHandle(window), "github")
                            .Where(item => item.Current.ControlType == ControlType.MenuItem &&
                                item.Current.ProcessId == context.Fixture.ProcessId && !item.Current.IsOffscreen));
                    }
                    var unique = matches.GroupBy(item => item.Current.AutomationId + ":" + item.Current.Name).ToArray();
                    if (unique.Length > 1 || unique.Length == 1 && unique[0].Count() > 1)
                        throw new InvalidOperationException("The chat GitHub menu item is ambiguous.");
                    if (unique.Length == 1) gitItem = unique[0].Single();
                    else Thread.Sleep(50);
                }
                if (gitItem == null) throw new InvalidOperationException("The exact chat GitHub menu item was not observed.");
                context.Record(new { Phase = "ChatGitItemObserved", AutomationId = gitItem.Current.AutomationId,
                    Name = gitItem.Current.Name, ProcessId = gitItem.Current.ProcessId });
            }

            internal void RequireNoGitModal()
            {
                if (FindGitModal().Length != 0) throw new InvalidOperationException("A Git modal existed before the chat action.");
            }

            internal void InvokeGitOnce()
            {
                try
                {
                    RequireSelectedScope();
                    context.Record(new { Phase = "ChatGitInvokeIntent", context.Label, CanonicalPath = context.Scope.Path,
                        ChatHandle = context.ChatHandle.ToInt64() });
                    context.ActionIssued = true;
                    context.Fixture.NativeExecutionUnsettled = true;
                    context.GitIntent.Set();
                    Pattern<InvokePattern>(gitItem, InvokePattern.Pattern).Invoke();
                    context.Record(new { Phase = "ChatGitInvokeReturned" });
                }
                catch (Exception error) { context.InvokerError = error; }
                finally { context.InvokerDone.Set(); }
            }

            private IntPtr[] FindGitModal()
            {
                var matches = new List<IntPtr>();
                foreach (IntPtr window in OwnedTopWindows())
                {
                    uint pid; uint tid = GetWindowThreadProcessId(window, out pid);
                    if (tid != context.Scope.ThreadId || !IsWindowVisible(window)) continue;
                    var title = new StringBuilder(128); GetWindowText(window, title, title.Capacity);
                    var cls = new StringBuilder(128); GetClassName(window, cls, cls.Capacity);
                    if (title.ToString() == "GitHub · VBAi" && cls.ToString().StartsWith("WindowsForms", StringComparison.Ordinal))
                        matches.Add(window);
                }
                return matches.ToArray();
            }

            internal void ObserveGitModal()
            {
                var watch = Stopwatch.StartNew();
                while (watch.ElapsedMilliseconds < 15000 && git == null)
                {
                    if (context.Stop) throw new InvalidOperationException("Coordinator stopped modal observation.");
                    var found = FindGitModal();
                    if (found.Length > 1) throw new InvalidOperationException("Ambiguous owned Word Git modals.");
                    if (found.Length == 0) { Thread.Sleep(50); continue; }
                    context.GitHandle = found[0]; Guard(context.GitHandle);
                    if (GetWindow(context.GitHandle, 4) != context.ChatHandle)
                        throw new InvalidOperationException("The Git modal is not owned by the selected Word chat window.");
                    git = AutomationElement.FromHandle(context.GitHandle);
                    if (git.Current.ProcessId != context.Fixture.ProcessId ||
                        unchecked((uint)git.Current.NativeWindowHandle) != unchecked((uint)context.GitHandle.ToInt64()))
                        throw new InvalidOperationException("The Word Git modal UIA identity differs from its native HWND.");
                    gitWindow = Pattern<WindowPattern>(git, WindowPattern.Pattern);
                }
                if (git == null) throw new TimeoutException("The chat Git action has no observed exact modal; no retry or cleanup.");
                context.ModalObserved = true;
                context.Record(new { Phase = "OwnedChatGitModalObserved", ProcessId = context.Fixture.ProcessId,
                    ThreadId = context.Scope.ThreadId, Handle = context.GitHandle.ToInt64(), Owner = context.ChatHandle.ToInt64() });
            }

            internal void CloseExactGitModal()
            {
                Guard(context.GitHandle);
                if (GetWindow(context.GitHandle, 4) != context.ChatHandle)
                    throw new InvalidOperationException("The exact chat Git modal owner changed before close.");
                context.Record(new { Phase = "ChatGitCloseIntent", Handle = context.GitHandle.ToInt64() });
                gitWindow.Close();
                var watch = Stopwatch.StartNew();
                while (watch.ElapsedMilliseconds < 10000)
                {
                    uint pid; uint tid = GetWindowThreadProcessId(context.GitHandle, out pid);
                    if (!IsWindow(context.GitHandle) && tid == 0)
                    {
                        context.Record(new { Phase = "ChatGitCloseObserved", Handle = context.GitHandle.ToInt64() });
                        return;
                    }
                    EmbeddedGitUiProtocol.RequireOwner(context.Fixture.ProcessId, context.Scope.ThreadId,
                        context.GitHandle.ToInt64(), (int)pid, tid, context.GitHandle.ToInt64());
                    Thread.Sleep(50);
                }
                throw new TimeoutException("The exact chat Git modal close outcome is uncertain.");
            }

            private static T Pattern<T>(AutomationElement item, AutomationPattern pattern) where T : class
            {
                object found;
                if (!item.TryGetCurrentPattern(pattern, out found) || !(found is T))
                    throw new InvalidOperationException("The required Word chat UIA pattern is absent: " + pattern.ProgrammaticName);
                return (T)found;
            }
        }
    }
}
