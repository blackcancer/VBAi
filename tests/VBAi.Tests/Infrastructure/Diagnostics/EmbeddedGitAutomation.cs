using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Automation;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Integration
{
    /// <summary>Targets only verified native leaf handles of the installed modal GitWindow.</summary>
    internal sealed class EmbeddedGitAutomation
    {
        private delegate bool Visitor(IntPtr hwnd, IntPtr data);
        [DllImport("user32.dll")] private static extern bool EnumWindows(Visitor visit, IntPtr data);
        [DllImport("user32.dll")] private static extern bool EnumChildWindows(IntPtr hwnd, Visitor visit, IntPtr data);
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint pid);
        [DllImport("user32.dll")] private static extern IntPtr GetWindow(IntPtr hwnd, uint command);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetClassName(IntPtr hwnd, StringBuilder text, int length);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetWindowText(IntPtr hwnd, StringBuilder text, int length);
        [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr hwnd);
        private readonly ExcelVbeFixture fixture;
        private readonly ExcelVbeFixture.EmbeddedGitScope scope;
        private readonly Action<object> record;
        private readonly Func<bool> stop;
        internal readonly EmbeddedGitUiProtocol Protocol;
        private IntPtr window;
        private AutomationElement root;
        private WindowPattern windowPattern;
        internal bool HasObservedWindow => windowPattern != null;
        private readonly HashSet<string> inventories = new HashSet<string>(StringComparer.Ordinal);
        internal EmbeddedGitAutomation(ExcelVbeFixture fixture, ExcelVbeFixture.EmbeddedGitScope scope, Action<object> record, Func<bool> stop)
        { this.fixture = fixture; this.scope = scope; this.record = record; this.stop = stop; Protocol = new EmbeddedGitUiProtocol(record); }

        internal void OpenedWindow()
        {
            var watch = Stopwatch.StartNew();
            while (window == IntPtr.Zero && watch.ElapsedMilliseconds < 15000)
            {
                if (stop()) throw new InvalidOperationException("Coordinator stopped discovery; no UI action is permitted.");
                fixture.RequireEmbeddedProcess(scope); var matches = new List<IntPtr>(); int count = 0;
                Visitor visitor = (hwnd, unused) => {
                    if (++count > 4096) return false;
                    uint pid; uint tid = GetWindowThreadProcessId(hwnd, out pid);
                    if (pid != fixture.ProcessId || tid != scope.ThreadId || GetWindow(hwnd, 4) != scope.VbeHandle) return true;
                    var text = new StringBuilder(256); GetWindowText(hwnd, text, text.Capacity);
                    var cls = new StringBuilder(256); GetClassName(hwnd, cls, cls.Capacity);
                    if (IsWindowVisible(hwnd) && text.ToString() == "GitHub · VBAi" && cls.ToString().StartsWith("WindowsForms", StringComparison.Ordinal)) matches.Add(hwnd);
                    return true;
                };
                if (!EnumWindows(visitor, IntPtr.Zero) || count > 4096) throw new InvalidOperationException("Bounded owned modal enumeration failed.");
                if (matches.Count > 1) throw new InvalidOperationException("Ambiguous owned GitWindow.");
                if (matches.Count == 1) window = matches[0]; else Thread.Sleep(50);
            }
            if (window == IntPtr.Zero) { Protocol.MarkUncertain("Modal menu has no observed GitWindow."); throw new TimeoutException("Owned embedded GitWindow did not become visible; no menu replay or cleanup."); }
            root = AutomationElement.FromHandle(window); Guard(root, window);
            windowPattern = Pattern<WindowPattern>(root, WindowPattern.Pattern);
            record(new { Phase = "OwnedModalObserved", ProcessId = fixture.ProcessId, ThreadId = scope.ThreadId,
                Handle = window.ToInt64(), Owner = scope.VbeHandle.ToInt64(), WindowPatternObserved = true });
        }

        internal void Link(string remote, string branch, string branchCommit)
        {
            Set("remote", remote); Set("branch", branch);
            string before = Text("status"); Invoke("connect");
            WaitTerminal("connect", before);
            string binding = Path.Combine(scope.Cache, "binding.json");
            Assert.IsTrue(File.Exists(binding), "A terminal connect error is a failure, not an idle pass.");
            var values = new System.Web.Script.Serialization.JavaScriptSerializer().Deserialize<Dictionary<string, string>>(File.ReadAllText(binding));
            Assert.AreEqual(remote, values["Remote"]); Assert.AreEqual(branch, values["Branch"]);
            using (var hash = System.Security.Cryptography.SHA256.Create())
            {
                string id = BitConverter.ToString(hash.ComputeHash(Encoding.UTF8.GetBytes(remote + "\n" + branch))).Replace("-", "");
                var repo = new MacroGitRepository(Path.Combine(scope.Cache, id + ".git"), branch, null);
                Assert.AreEqual(branchCommit, repo.Resolve("refs/remotes/origin/selected"), "Only the explicitly authorized synthetic branch revision is accepted.");
            }
        }

        internal void CompareOnce()
        {
            // After checkpoint_create, the terminal Compare status differs even if busy completed between observations.
            string before = Text("status"); Invoke("compare"); WaitTerminal("compare", before);
        }

        internal void Checkpoint(string remote, string branch, string nonce, string tabName)
        {
            var tabs = Leaf("tabs"); var items = tabs.FindAll(TreeScope.Children,
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.TabItem)).Cast<AutomationElement>().ToArray();
            if (items.Length < 1 || items.Length > 12) throw new InvalidOperationException("Bounded tab item inventory failed.");
            var matches = items.Where(item => item.Current.ProcessId == fixture.ProcessId && item.Current.Name == tabName).ToArray();
            if (matches.Length != 1) throw new InvalidOperationException("The exact observed Checkpoints tab is absent or ambiguous.");
            var selection = Pattern<SelectionItemPattern>(matches[0], SelectionItemPattern.Pattern);
            Guard(tabs, new IntPtr(tabs.Current.NativeWindowHandle));
            Protocol.EmitOnce("select-checkpoints", selection.Select);
            Assert.IsTrue(selection.Current.IsSelected); Protocol.Terminal("select-checkpoints", true, true);
            Set("checkpointName", nonce);
            string before = Text("status"); Invoke("checkpointCreate"); WaitTerminal("checkpointCreate", before);
            using (var hash = System.Security.Cryptography.SHA256.Create())
            {
                string id = BitConverter.ToString(hash.ComputeHash(Encoding.UTF8.GetBytes(remote + "\n" + branch))).Replace("-", "");
                var repo = new MacroGitRepository(Path.Combine(scope.Cache, id + ".git"), branch, null);
                var checkpoint = repo.Checkpoints().Single(item => item.Label == nonce);
                Assert.AreEqual(checkpoint.Commit, repo.Resolve("refs/codex/checkpoints/" + checkpoint.Id));
                EmbeddedGitSnapshotOracle.Verify(scope.Baseline, repo.Read(checkpoint.Commit));
                record(new { Phase = "CheckpointContentVerified", checkpoint.Id, checkpoint.Commit, checkpoint.Label,
                    Files = EmbeddedGitSnapshotOracle.Describe(repo.Read(checkpoint.Commit)),
                    Scope = "Installed owner-dispatched capture and local checkpoint only; no remote publish/import/reopen acceptance" });
            }
        }

        internal void CloseKnownTerminal()
        {
            if (!Protocol.CanClose) throw new InvalidOperationException("Pending or uncertain UI delivery forbids modal Close.");
            Guard(root, window);
            record(new { Phase = "WindowClosePatternObserved", Pattern = WindowPattern.Pattern.ProgrammaticName });
            Protocol.EmitOnce("window-close", windowPattern.Close);
            var watch = Stopwatch.StartNew();
            while (watch.ElapsedMilliseconds < 10000)
            {
                uint pid; uint tid = GetWindowThreadProcessId(window, out pid);
                if (tid == 0 && pid == 0) { Protocol.Terminal("window-close", true, true); return; }
                EmbeddedGitUiProtocol.RequireOwner(fixture.ProcessId, scope.ThreadId, window.ToInt64(), (int)pid, tid, window.ToInt64());
                Thread.Sleep(50);
            }
            Protocol.MarkUncertain("Exact modal destruction was not observed.");
            throw new TimeoutException("Modal Close outcome is uncertain; no second Close or host cleanup.");
        }

        private void Set(string id, string value)
        {
            var item = Leaf(id); var pattern = Pattern<ValuePattern>(item, ValuePattern.Pattern);
            Assert.IsFalse(pattern.Current.IsReadOnly); RequireInteractive(item);
            Protocol.EmitOnce("value-" + id, () => pattern.SetValue(value));
            Assert.AreEqual(value, pattern.Current.Value); Protocol.Terminal("value-" + id, true, true);
        }
        private void Invoke(string id)
        {
            var item = Leaf(id); var pattern = Pattern<InvokePattern>(item, InvokePattern.Pattern); RequireInteractive(item);
            Protocol.EmitOnce(id, pattern.Invoke);
        }
        private void WaitTerminal(string id, string before)
        {
            bool busy = false; var watch = Stopwatch.StartNew();
            while (watch.ElapsedMilliseconds < 60000)
            {
                string text = Text("status"); bool idle = Leaf("compare").Current.IsEnabled || Leaf("connect").Current.IsEnabled;
                busy |= !idle;
                if (EmbeddedGitUiProtocol.IsTerminal(idle, busy, before, text))
                {
                    record(new { Phase = "UiTerminalObserved", Action = id, Status = text, BusyObserved = busy, Idle = idle });
                    int classification = ClassifyTerminal(id, text);
                    bool error = classification < 0;
                    if (classification == 0) throw new InvalidOperationException("Unclassified UI terminal status; preserve rather than infer success or close.");
                    Protocol.Terminal(id, true, !error);
                    if (error) throw new InvalidOperationException("Embedded Git returned a known terminal error: " + text);
                    return;
                }
                Thread.Sleep(50);
            }
            Protocol.MarkUncertain("No terminal operation evidence before bounded deadline.");
            throw new TimeoutException("Embedded Git operation remains uncertain; no replay or automatic Close/Quit.");
        }
        internal static int ClassifyTerminal(string id, string text)
        {
            if (string.IsNullOrEmpty(text)) return 0;
            if (Labels("Check the connection, account and Git state, then retry.").Any(x => text.EndsWith(x, StringComparison.Ordinal))) return -1;
            bool success = id == "checkpointCreate"
                ? Labels("Operation complete: ").Any(x => text == x + "checkpoint_create")
                : (id == "connect" || id == "compare") &&
                    (Labels("First link: commit then push to publish, or pull to import the repository with a backup first.").Contains(text)
                    || Labels("VBA matches the last synchronized state.").Contains(text)
                    || Labels(" file(s) changed since the last synchronization.").Any(x => text.EndsWith(x, StringComparison.Ordinal) && char.IsDigit(text[0])));
            return success ? 1 : 0;
        }
        private string Text(string id) => Leaf(id).Current.Name;
        private static string[] Labels(string english)
        {
            return UiLanguages.All.Select(language => new System.Resources.ResourceManager("VBAi.Localization.UiStrings" + language.ResourceSuffix,
                typeof(VbeSession).Assembly).GetString(english, System.Globalization.CultureInfo.InvariantCulture) ?? english)
                .Concat(new[] { english }).Distinct(StringComparer.Ordinal).ToArray();
        }
        private void RequireInteractive(AutomationElement item)
        { if (!item.Current.IsEnabled || item.Current.IsOffscreen) throw new InvalidOperationException("The exact UI control is not interactive."); }
        private static T Pattern<T>(AutomationElement item, AutomationPattern pattern) where T : class
        {
            object found; if (!item.TryGetCurrentPattern(pattern, out found) || !(found is T))
                throw new InvalidOperationException("Required observed UIA pattern is absent: " + pattern.ProgrammaticName);
            return (T)found;
        }
        private void Guard(AutomationElement item, IntPtr hwnd)
        {
            if (stop()) throw new InvalidOperationException("Coordinator deadline forbids any further UI action.");
            fixture.RequireEmbeddedProcess(scope); uint pid; uint tid = GetWindowThreadProcessId(hwnd, out pid);
            EmbeddedGitUiProtocol.RequireOwner(fixture.ProcessId, scope.ThreadId, hwnd.ToInt64(), (int)pid, tid, hwnd.ToInt64());
            if (item.Current.ProcessId != fixture.ProcessId || unchecked((uint)item.Current.NativeWindowHandle) != unchecked((uint)hwnd.ToInt64()))
                throw new InvalidOperationException("UIA provider identity differs from the exact native HWND.");
            if (window != IntPtr.Zero && GetWindow(window, 4) != scope.VbeHandle) throw new InvalidOperationException("Owned modal parent changed.");
        }
        private AutomationElement Leaf(string id)
        {
            Guard(root, window); var handles = new List<IntPtr>();
            Visitor visit = (hwnd, unused) => { handles.Add(hwnd); return handles.Count < 500; };
            if (!EnumChildWindows(window, visit, IntPtr.Zero) || handles.Count >= 500) throw new InvalidOperationException("Bounded native child enumeration failed.");
            var matches = new List<AutomationElement>();
            foreach (var hwnd in handles)
            {
                var cls = new StringBuilder(256); if (GetClassName(hwnd, cls, cls.Capacity) == 0) throw new InvalidOperationException("Native child class unreadable.");
                string name = cls.ToString();
                if (!name.Contains(".EDIT.") && !name.Contains(".BUTTON.") && !name.Contains(".STATIC.") && !name.Contains("SysTabControl32") && !name.Contains(".LISTBOX.")) continue;
                var item = AutomationElement.FromHandle(hwnd); Guard(item, hwnd);
                if (item.Current.AutomationId == id) matches.Add(item);
            }
            if (matches.Count != 1) throw new InvalidOperationException("Exact native leaf absent or ambiguous: " + id);
            var result = matches[0];
            if (inventories.Add(id)) record(new { Phase = "PatternInventory", Id = id, Handle = result.Current.NativeWindowHandle,
                Type = result.Current.ControlType.ProgrammaticName, Patterns = result.GetSupportedPatterns().Select(x => x.ProgrammaticName).ToArray() });
            return result;
        }
    }
}
