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
                        catch (Exception cleanup)
                        {
                            context.OwnerError = context.OwnerError == null ? cleanup : new AggregateException(context.OwnerError, cleanup);
                            if (context.Fixture.WordGitMustRetain)
                            { context.Retain = true; lock (Retained) if (!Retained.Contains(context)) Retained.Add(context); }
                        }
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
                automation.CaptureModalOwner();
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
            [DllImport("user32.dll", SetLastError = true)] private static extern bool EnumWindows(WordChatTopWindowInventory.Visitor visitor, IntPtr data);
            [DllImport("user32.dll")] private static extern bool EnumChildWindows(IntPtr parent, WordChatWindowDiscovery.NativeVisitor visitor, IntPtr data);
            [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr window, out uint pid);
            [DllImport("user32.dll")] private static extern bool IsChild(IntPtr parent, IntPtr child);
            [DllImport("user32.dll")] private static extern bool IsWindow(IntPtr window);
            [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr window);
            [DllImport("user32.dll")] private static extern IntPtr GetWindow(IntPtr window, uint command);
            [DllImport("user32.dll")] private static extern IntPtr GetAncestor(IntPtr window, uint flags);
            [DllImport("user32.dll")] private static extern IntPtr GetParent(IntPtr window);
            [DllImport("user32.dll", EntryPoint = "GetWindowLongPtrW")] private static extern IntPtr GetWindowLongPtr(IntPtr window, int index);
            [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetClassName(IntPtr window, StringBuilder text, int capacity);
            [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetWindowText(IntPtr window, StringBuilder text, int capacity);
            private readonly Context context;
            private AutomationElement chat, scopePicker, gitItem, git;
            private IntPtr gitPopupHandle;
            private WordChatGitMenuDiscovery.Candidate selectedPopup;
            private SelectionPattern scopeSelection;
            private WindowPattern gitWindow;
            private WordChatWindowDiscovery.OwnerIdentity modalOwner;
            internal WordChatGitAutomation(Context context) { this.context = context; }

            private IntPtr[] OwnedTopWindows()
            {
                return WordChatTopWindowInventory.Read(EnumWindows, window => {
                    uint pid; uint tid = GetWindowThreadProcessId(window, out pid);
                    return new WordChatTopWindowInventory.Identity {
                        ProcessId = (int)pid, ThreadId = tid, Visible = IsWindowVisible(window)
                    };
                }, context.Fixture.ProcessId, context.Scope.ThreadId, result =>
                    context.Record(new { Phase = "WordTopWindowInventoryRefused", result.ApiReturned,
                        result.VisitedTotal, result.OwnedProcessCount, result.ExactThreadVisibleCount,
                        result.GlobalBoundHit, result.FailureStatus,
                        NativeLastError = !result.ApiReturned && !result.GlobalBoundHit
                            ? (int?)Marshal.GetLastWin32Error() : null }));
            }

            private static string NativeClass(IntPtr window)
            {
                var value = new StringBuilder(128);
                GetClassName(window, value, value.Capacity);
                return value.ToString();
            }

            private static WordChatGitMenuDiscovery.OwnerShape ReadPopupOwner(IntPtr handle)
            {
                uint pid; uint tid = GetWindowThreadProcessId(handle, out pid);
                return new WordChatGitMenuDiscovery.OwnerShape {
                    Handle = handle.ToInt64(), Live = IsWindow(handle), ProcessId = (int)pid, ThreadId = tid,
                    ClassName = NativeClass(handle), Visible = IsWindowVisible(handle),
                    Parent = GetParent(handle).ToInt64(), Root = GetAncestor(handle, 2).ToInt64(),
                    Owner = GetWindow(handle, 4).ToInt64(),
                    Style = unchecked((uint)GetWindowLongPtr(handle, -16).ToInt64()),
                    ExStyle = unchecked((uint)GetWindowLongPtr(handle, -20).ToInt64())
                };
            }

            private static object OwnerEvidence(WordChatGitMenuDiscovery.OwnerShape owner)
                => owner == null ? null : new {
                    owner.Handle, owner.Live, owner.ProcessId, owner.ThreadId, owner.ClassName,
                    owner.Visible, owner.Parent, owner.Root, owner.Owner, owner.Style, owner.ExStyle
                };

            private static AutomationElement[] Descendants(AutomationElement root, string id)
                => root.FindAll(TreeScope.Descendants, new PropertyCondition(AutomationElement.AutomationIdProperty, id))
                    .Cast<AutomationElement>().ToArray();

            private static IntPtr UiHandle(int handle) => new IntPtr(unchecked((long)(uint)handle));

            private static IntPtr NativeAncestorHandle(AutomationElement element)
            {
                for (int depth = 0; element != null && depth < 16; depth++)
                {
                    IntPtr handle = UiHandle(element.Current.NativeWindowHandle);
                    if (handle != IntPtr.Zero) return handle;
                    element = TreeWalker.RawViewWalker.GetParent(element);
                }
                return IntPtr.Zero;
            }

            private IntPtr[] NativeChildren(IntPtr parent)
            {
                Guard(parent);
                var found = WordChatWindowDiscovery.ReadNativeChildren(parent, EnumChildWindows);
                Guard(parent);
                return found;
            }

            private WordChatWindowDiscovery.Candidate DescribeChatCandidate(IntPtr window, string nativeClass,
                AutomationElement element)
            {
                uint nativePid, pickerPid = 0, optionsPid = 0;
                uint nativeTid = GetWindowThreadProcessId(window, out nativePid);
                var pickers = Descendants(element, "scopePicker");
                var options = Descendants(element, "options");
                IntPtr pickerHandle = pickers.Length == 1 ? NativeAncestorHandle(pickers[0]) : IntPtr.Zero;
                IntPtr optionsHandle = options.Length == 1 ? NativeAncestorHandle(options[0]) : IntPtr.Zero;
                uint pickerTid = pickerHandle == IntPtr.Zero ? 0 : GetWindowThreadProcessId(pickerHandle, out pickerPid);
                uint optionsTid = optionsHandle == IntPtr.Zero ? 0 : GetWindowThreadProcessId(optionsHandle, out optionsPid);
                return new WordChatWindowDiscovery.Candidate {
                    Handle = window.ToInt64(), NativeProcessId = (int)nativePid, NativeThreadId = nativeTid,
                    UiProcessId = element.Current.ProcessId, Visible = IsWindowVisible(window),
                    WithinOwnedVbe = IsChild(context.Scope.VbeHandle, window) || GetWindow(window, 4) == context.Scope.VbeHandle,
                    FixedChatCaption = true,
                    NativeClass = nativeClass, ControlType = element.Current.ControlType.ProgrammaticName,
                    ScopePickerCount = pickers.Length, OptionsCount = options.Length,
                    ScopePickerProcessId = pickers.Length == 1 ? pickers[0].Current.ProcessId : 0,
                    OptionsProcessId = options.Length == 1 ? options[0].Current.ProcessId : 0,
                    ScopePickerType = pickers.Length == 1 ? pickers[0].Current.ControlType.ProgrammaticName : null,
                    OptionsType = options.Length == 1 ? options[0].Current.ControlType.ProgrammaticName : null,
                    ScopePickerHandle = pickerHandle.ToInt64(), OptionsHandle = optionsHandle.ToInt64(),
                    ScopePickerThreadId = pickerTid, OptionsThreadId = optionsTid,
                    ScopePickerWithinChat = pickerHandle != IntPtr.Zero &&
                        (pickerHandle == window || IsChild(window, pickerHandle)),
                    OptionsWithinChat = optionsHandle != IntPtr.Zero &&
                        (optionsHandle == window || IsChild(window, optionsHandle))
                };
            }

            private AutomationElement DiscoverChatWindow()
            {
                Guard(context.Scope.VbeHandle);
                var handles = new HashSet<IntPtr>(NativeChildren(context.Scope.VbeHandle));
                foreach (IntPtr top in OwnedTopWindows().Where(top => GetWindow(top, 4) == context.Scope.VbeHandle))
                {
                    uint pid; uint tid = GetWindowThreadProcessId(top, out pid);
                    if (pid != context.Fixture.ProcessId || tid != context.Scope.ThreadId || !IsWindowVisible(top)) continue;
                    handles.Add(top);
                    foreach (IntPtr child in NativeChildren(top)) handles.Add(child);
                }
                if (handles.Count > 4096) throw new InvalidOperationException("Bounded Word chat native-window inventory failed.");
                var candidates = new List<WordChatWindowDiscovery.Candidate>();
                var elements = new Dictionary<long, AutomationElement>();
                int exactThreadWindows = 0, windowsFormsCount = 0, fixedCaptionCount = 0;
                var controlTypes = new Dictionary<string, int>(StringComparer.Ordinal);
                try
                {
                    foreach (IntPtr window in handles)
                    {
                        uint pid; uint tid = GetWindowThreadProcessId(window, out pid);
                        if (pid != context.Fixture.ProcessId || tid != context.Scope.ThreadId || !IsWindowVisible(window)) continue;
                        exactThreadWindows++;
                        var cls = new StringBuilder(128);
                        if (GetClassName(window, cls, cls.Capacity) == 0)
                            throw new InvalidOperationException("Owned Word child-window class could not be read.");
                        if (!cls.ToString().StartsWith("WindowsForms", StringComparison.Ordinal)) continue;
                        windowsFormsCount++;
                        var element = AutomationElement.FromHandle(window);
                        string type = element.Current.ControlType.ProgrammaticName;
                        controlTypes[type] = controlTypes.TryGetValue(type, out int count) ? count + 1 : 1;
                        if (element.Current.ControlType != ControlType.Window &&
                            element.Current.ControlType != ControlType.Pane) continue;
                        var title = new StringBuilder(128);
                        GetWindowText(window, title, title.Capacity);
                        if (!string.Equals(title.ToString(), "VBAi — Your AI agent for VBA", StringComparison.Ordinal)) continue;
                        fixedCaptionCount++;
                        if (candidates.Count >= 64) throw new InvalidOperationException("Bounded Word chat Form candidate inventory exceeded.");
                        var candidate = DescribeChatCandidate(window, cls.ToString(), element);
                        candidates.Add(candidate); elements.Add(candidate.Handle, element);
                    }
                }
                finally
                {
                    // No captions, transcript text, paths or provider state enter this pre-action receipt.
                    context.Record(new { Phase = "ChatNativeWindowInventory", VbeHandle = context.Scope.VbeHandle.ToInt64(),
                        ProcessId = context.Fixture.ProcessId, ThreadId = context.Scope.ThreadId,
                        NativeWindowCount = handles.Count, ExactThreadVisibleCount = exactThreadWindows,
                        WindowsFormsCount = windowsFormsCount, ControlTypes = controlTypes,
                        FixedChatCaptionCount = fixedCaptionCount,
                        Candidates = candidates.Select(item => new { item.Handle, item.NativeClass, item.ControlType,
                            item.NativeProcessId, item.NativeThreadId, item.UiProcessId, item.Visible,
                            item.WithinOwnedVbe, item.FixedChatCaption,
                            item.ScopePickerCount, item.OptionsCount, item.ScopePickerType, item.OptionsType,
                            item.ScopePickerHandle, item.OptionsHandle, item.ScopePickerThreadId, item.OptionsThreadId,
                            item.ScopePickerProcessId, item.OptionsProcessId,
                            item.ScopePickerWithinChat, item.OptionsWithinChat }).ToArray() });
                }
                return elements[WordChatWindowDiscovery.RequireUnique(candidates, context.Fixture.ProcessId,
                    context.Scope.ThreadId).Handle];
            }

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
                chat = DiscoverChatWindow(); context.ChatHandle = new IntPtr(chat.Current.NativeWindowHandle);
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

            private WordChatWindowDiscovery.OwnerIdentity ReadModalOwner()
            {
                Guard(context.ChatHandle);
                IntPtr vbeRoot = GetAncestor(context.Scope.VbeHandle, 2); // GA_ROOT follows native parents, not owners.
                IntPtr chatRoot = GetAncestor(context.ChatHandle, 2);
                uint vbePid, chatPid;
                uint vbeTid = GetWindowThreadProcessId(vbeRoot, out vbePid);
                uint chatTid = GetWindowThreadProcessId(chatRoot, out chatPid);
                return new WordChatWindowDiscovery.OwnerIdentity {
                    VbeHandle = context.Scope.VbeHandle.ToInt64(), VbeRoot = vbeRoot.ToInt64(),
                    ChatHandle = context.ChatHandle.ToInt64(), ChatRoot = chatRoot.ToInt64(),
                    ChatOwner = GetWindow(context.ChatHandle, 4).ToInt64(),
                    ChatWithinVbe = IsChild(context.Scope.VbeHandle, context.ChatHandle),
                    VbeRootProcessId = (int)vbePid, ChatRootProcessId = (int)chatPid,
                    VbeRootThreadId = vbeTid, ChatRootThreadId = chatTid
                };
            }

            internal void CaptureModalOwner()
            {
                RequireSelectedScope();
                modalOwner = ReadModalOwner();
                long expected = WordChatWindowDiscovery.RequireModalOwner(modalOwner,
                    context.Fixture.ProcessId, context.Scope.ThreadId);
                context.Record(new { Phase = "ChatModalOwnerPreflight", modalOwner.VbeHandle, modalOwner.VbeRoot,
                    modalOwner.ChatHandle, modalOwner.ChatRoot, modalOwner.ChatOwner,
                    modalOwner.ChatWithinVbe, ExpectedModalOwner = expected,
                    modalOwner.VbeRootProcessId, modalOwner.VbeRootThreadId,
                    modalOwner.ChatRootProcessId, modalOwner.ChatRootThreadId });
            }

            private void RequireSameModalOwner(IntPtr observedOwner)
            {
                WordChatWindowDiscovery.RequireUnchangedModalOwner(modalOwner, ReadModalOwner(), observedOwner.ToInt64(),
                    context.Fixture.ProcessId, context.Scope.ThreadId);
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
                long expectedOwner = WordChatWindowDiscovery.RequireModalOwner(modalOwner,
                    context.Fixture.ProcessId, context.Scope.ThreadId);
                var visibleBefore = new HashSet<IntPtr>(OwnedTopWindows().Where(window => {
                    uint pid; uint tid = GetWindowThreadProcessId(window, out pid);
                    return pid == context.Fixture.ProcessId && tid == context.Scope.ThreadId && IsWindowVisible(window);
                }));
                context.Record(new { Phase = "ChatOptionsIntent", ChatHandle = context.ChatHandle.ToInt64() });
                if (PrivateDesktopUiAction.Enabled)
                {
                    IntPtr optionsHandle = NativeAncestorHandle(buttons[0]);
                    if (optionsHandle == IntPtr.Zero ||
                        unchecked((long)(uint)buttons[0].Current.NativeWindowHandle) != optionsHandle.ToInt64())
                        throw new InvalidOperationException("The exact Word chat Options button has no native leaf HWND.");
                    context.Fixture.NativeExecutionUnsettled = true;
                    PrivateDesktopUiAction.ClickButtonOnce(buttons[0], "options", buttons[0].Current.Name,
                        context.ChatHandle, optionsHandle, context.Fixture.ProcessId, context.Scope.ThreadId);
                }
                else Pattern<InvokePattern>(buttons[0], InvokePattern.Pattern).Invoke();
                var watch = Stopwatch.StartNew();
                var last = new List<WordChatGitMenuDiscovery.Candidate>();
                int nativeVisible = 0;
                try
                {
                    while (watch.ElapsedMilliseconds < 5000 && gitItem == null)
                    {
                        last = new List<WordChatGitMenuDiscovery.Candidate>();
                        var exactItems = new Dictionary<long, AutomationElement>();
                        nativeVisible = 0;
                        foreach (IntPtr window in OwnedTopWindows())
                        {
                            uint pid; uint tid = GetWindowThreadProcessId(window, out pid);
                            if (pid != context.Fixture.ProcessId || tid != context.Scope.ThreadId ||
                                !IsWindowVisible(window)) continue;
                            nativeVisible++;
                            string cls = NativeClass(window);
                            if (!cls.StartsWith("WindowsForms", StringComparison.Ordinal)) continue;
                            var popup = AutomationElement.FromHandle(window);
                            if (popup.Current.ControlType != ControlType.Menu) continue;
                            var items = popup.FindAll(TreeScope.Descendants,
                                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.MenuItem))
                                .Cast<AutomationElement>().ToArray();
                            string label = UiText.Get("GitHub · synchronize VBA…");
                            var matching = items.Where(item => WordChatGitMenuDiscovery.IsExactGitItem(
                                item.Current.Name, label, item.Current.ControlType.ProgrammaticName,
                                item.Current.ProcessId, context.Fixture.ProcessId)).ToArray();
                            var enabled = matching.Where(item => item.Current.IsEnabled && !item.Current.IsOffscreen).ToArray();
                            var candidate = new WordChatGitMenuDiscovery.Candidate {
                                PopupHandle = window.ToInt64(), OwnerHandle = GetWindow(window, 4).ToInt64(),
                                NativeProcessId = (int)pid, NativeThreadId = tid, UiProcessId = popup.Current.ProcessId,
                                Visible = true, NewlyVisible = !visibleBefore.Contains(window),
                                NativeClass = cls, UiType = popup.Current.ControlType.ProgrammaticName,
                                MenuItemCount = items.Length, GitLabelMatches = matching.Length,
                                EnabledGitMatches = enabled.Length,
                                GitItemProcessId = matching.Length == 1 ? matching[0].Current.ProcessId : 0,
                                GitItemNativeAncestor = matching.Length == 1
                                    ? NativeAncestorHandle(matching[0]).ToInt64() : 0
                            };
                            candidate.OwnerShape = ReadPopupOwner(new IntPtr(candidate.OwnerHandle));
                            last.Add(candidate);
                            if (enabled.Length == 1) exactItems.Add(window.ToInt64(), enabled[0]);
                        }
                        var eligible = last.Where(item => WordChatGitMenuDiscovery.HasStrictPopupOwner(item,
                                context.Fixture.ProcessId, context.Scope.ThreadId, expectedOwner) && item.NewlyVisible &&
                            item.GitLabelMatches == 1 && item.EnabledGitMatches == 1).ToArray();
                        if (eligible.Length > 0)
                        {
                            var selected = WordChatGitMenuDiscovery.RequireUnique(last, context.Fixture.ProcessId,
                                context.Scope.ThreadId, expectedOwner);
                            gitPopupHandle = new IntPtr(selected.PopupHandle);
                            selectedPopup = selected;
                            gitItem = exactItems[selected.PopupHandle];
                        }
                        else Thread.Sleep(50);
                    }
                }
                finally
                {
                    // The popup receipt contains only native identity and product-label match counts.
                    context.Record(new { Phase = "ChatOptionsPopupInventory", ProcessId = context.Fixture.ProcessId,
                        ThreadId = context.Scope.ThreadId, ExpectedOwner = expectedOwner,
                        VisibleTopWindowCount = nativeVisible, Candidates = last.Select(item => new {
                            item.PopupHandle, item.OwnerHandle, item.NativeProcessId, item.NativeThreadId,
                            item.UiProcessId, item.NativeClass, item.UiType, item.Visible, item.NewlyVisible,
                            item.MenuItemCount, item.GitLabelMatches, item.EnabledGitMatches,
                            item.GitItemProcessId, item.GitItemNativeAncestor,
                            Owner = OwnerEvidence(item.OwnerShape) }).ToArray() });
                }
                if (gitItem == null) throw new InvalidOperationException("The exact chat GitHub menu item was not observed.");
                context.Record(new { Phase = "ChatGitItemObserved", PopupHandle = gitPopupHandle.ToInt64(),
                    LocalizedProductLabelMatched = true, ProcessId = gitItem.Current.ProcessId,
                    NativeAncestorHandle = NativeAncestorHandle(gitItem).ToInt64(),
                    PopupNativeClass = selectedPopup.NativeClass,
                    PopupOwnerHandle = selectedPopup.OwnerHandle,
                    PopupOwner = OwnerEvidence(selectedPopup.OwnerShape) });
                if (PrivateDesktopUiAction.Enabled) context.Fixture.NativeExecutionUnsettled = false;
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
                    RequireSameModalOwner(new IntPtr(WordChatWindowDiscovery.RequireModalOwner(modalOwner,
                        context.Fixture.ProcessId, context.Scope.ThreadId)));
                    Guard(gitPopupHandle);
                    long expectedRoot = WordChatWindowDiscovery.RequireModalOwner(modalOwner,
                        context.Fixture.ProcessId, context.Scope.ThreadId);
                    long currentPopupOwner = GetWindow(gitPopupHandle, 4).ToInt64();
                    string currentPopupClass = NativeClass(gitPopupHandle);
                    var currentOwner = ReadPopupOwner(new IntPtr(currentPopupOwner));
                    WordChatGitMenuDiscovery.RequireUnchangedPopupOwner(selectedPopup, currentPopupClass,
                        currentPopupOwner, currentOwner,
                        context.Fixture.ProcessId, context.Scope.ThreadId, expectedRoot);
                    if (!IsWindowVisible(gitPopupHandle) ||
                        AutomationElement.FromHandle(gitPopupHandle).Current.ControlType != ControlType.Menu ||
                        !WordChatGitMenuDiscovery.IsExactGitItem(gitItem.Current.Name,
                            UiText.Get("GitHub · synchronize VBA…"), gitItem.Current.ControlType.ProgrammaticName,
                            gitItem.Current.ProcessId, context.Fixture.ProcessId) ||
                        NativeAncestorHandle(gitItem) != gitPopupHandle ||
                        !gitItem.Current.IsEnabled || gitItem.Current.IsOffscreen)
                        throw new InvalidOperationException("The exact localized Word chat Git menu item changed before invocation.");
                    context.Record(new { Phase = "ChatOptionsOwnerRevalidated", PopupHandle = gitPopupHandle.ToInt64(),
                        PopupNativeClass = currentPopupClass, PopupOwnerHandle = currentPopupOwner,
                        PopupOwner = OwnerEvidence(currentOwner), GitItemNativeAncestor = NativeAncestorHandle(gitItem).ToInt64() });
                    context.Record(new { Phase = "ChatGitInvokeIntent", context.Label, CanonicalPath = context.Scope.Path,
                        ChatHandle = context.ChatHandle.ToInt64() });
                    context.ActionIssued = true;
                    context.Fixture.NativeExecutionUnsettled = true;
                    context.GitIntent.Set();
                    if (PrivateDesktopUiAction.Enabled)
                        PrivateDesktopUiAction.InvokeVirtualGitOnce(gitItem, gitPopupHandle,
                            UiText.Get("GitHub · synchronize VBA…"), context.Fixture.ProcessId, context.Scope.ThreadId);
                    else Pattern<InvokePattern>(gitItem, InvokePattern.Pattern).Invoke();
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
                    IntPtr actualOwner = GetWindow(context.GitHandle, 4);
                    RequireSameModalOwner(actualOwner);
                    git = AutomationElement.FromHandle(context.GitHandle);
                    if (git.Current.ProcessId != context.Fixture.ProcessId ||
                        unchecked((uint)git.Current.NativeWindowHandle) != unchecked((uint)context.GitHandle.ToInt64()))
                        throw new InvalidOperationException("The Word Git modal UIA identity differs from its native HWND.");
                    gitWindow = Pattern<WindowPattern>(git, WindowPattern.Pattern);
                }
                if (git == null) throw new TimeoutException("The chat Git action has no observed exact modal; no retry or cleanup.");
                context.ModalObserved = true;
                context.Record(new { Phase = "OwnedChatGitModalObserved", ProcessId = context.Fixture.ProcessId,
                    ThreadId = context.Scope.ThreadId, Handle = context.GitHandle.ToInt64(),
                    Owner = GetWindow(context.GitHandle, 4).ToInt64(), ChatHandle = context.ChatHandle.ToInt64() });
            }

            internal void CloseExactGitModal()
            {
                Guard(context.GitHandle);
                RequireSameModalOwner(GetWindow(context.GitHandle, 4));
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
