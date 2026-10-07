using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Resources;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Automation;

namespace VBAi.Tests.Integration
{
    internal sealed partial class ExcelVbeFixture
    {
        internal const string MonacoSynchronizedStatus = "Synchronized with VBA. Save the macro in its host application.";
        internal const string MonacoClosedStatus = "The VBA project is closed. Your draft is preserved.";
        private object monacoLiveWorkbook;
        private string monacoClosedPath, monacoLivePath;
        private bool monacoPrimaryCloseInvoked, monacoLiveCloseInvoked;
        private delegate bool MonacoChildCallback(IntPtr child, IntPtr parameter);
        [DllImport("user32.dll")] private static extern bool EnumChildWindows(IntPtr parent, MonacoChildCallback callback, IntPtr parameter);
        [DllImport("user32.dll")] private static extern bool IsChild(IntPtr parent, IntPtr child);
        [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr window);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetWindowText(IntPtr window, StringBuilder text, int maximum);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetClassName(IntPtr window, StringBuilder text, int maximum);

        /// <summary>Scalar read-only observations from the installed child editor, never a detached testhost form.</summary>
        internal sealed class MonacoNativeObservation
        {
            public long VbeHandle { get; set; }
            public int VbeProcessId { get; set; }
            public long EditorHandle { get; set; }
            public int EditorProcessId { get; set; }
            public bool Embedded { get; set; }
            public bool Visible { get; set; }
            public string EditorClass { get; set; }
            public string EditorCaption { get; set; }
            public int StatusProcessId { get; set; }
            public string StatusAutomationId { get; set; }
            public string StatusClassName { get; set; }
            public string StatusText { get; set; }
            public bool StatusVisible { get; set; }
            public int StatusCount { get; set; }
            public string[] Tabs { get; set; } = new string[0];
            public string[] SelectedTabs { get; set; } = new string[0];
            public int[] TabProcessIds { get; set; } = new int[0];
            public string RenderedText { get; set; }
            public int[] TextProviderProcessIds { get; set; } = new int[0];
            public int ObservedElements { get; set; }
            public bool TreeTruncated { get; set; }
            public string ObservationError { get; set; }
        }

        internal static void RequireMonacoCandidate(Guid expected, Guid testReference, int pid, IDictionary<string, object> status)
        {
            Assert.AreNotEqual(Guid.Empty, expected); Assert.IsTrue(pid > 0);
            Assert.AreEqual(expected, testReference, "The isolated test output must reference the frozen installed payload.");
            Assert.AreEqual(pid, Convert.ToInt32(status["HostProcessId"]));
            Assert.AreEqual(expected.ToString("D"), Convert.ToString(status["AssemblyModuleVersionId"]), true);
        }

        internal void PreserveMonacoNativeOutcome()
        {
            PreserveForDiagnosticRecovery = true;
            lock (retainedBootstraps) if (!retainedBootstraps.Contains(this)) retainedBootstraps.Add(this);
        }

        private void RequireMonacoOwner()
        {
            Assert.IsTrue(owned && ownedProcess != null && !ownedProcess.HasExited, "The original owned Excel handle must remain alive.");
            uint pid; GetWindowThreadProcessId(new IntPtr(Convert.ToInt64(((dynamic)application).Hwnd)), out pid);
            Assert.AreEqual((uint)ProcessId, pid, "Native Excel ownership changed; do not rebind or clean up another process.");
        }

        private void RequireMonacoBook(object book, string path)
        {
            RequireMonacoOwner(); Assert.IsNotNull(book);
            string expected = Path.GetFullPath(path);
            Assert.IsTrue(expected.StartsWith(Path.GetFullPath(Root) + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase));
            Assert.AreEqual(expected, Path.GetFullPath(Convert.ToString(((dynamic)book).FullName)), true);
            object project = null;
            try { project = ((dynamic)book).VBProject; Assert.AreEqual(2, Convert.ToInt32(((dynamic)project).Mode)); }
            finally { BalanceMonacoGetter(project); }
        }

        internal void PrepareMonacoBooks(string closedPath, string livePath)
        {
            RequireMonacoOwner(); Assert.IsNull(monacoLiveWorkbook);
            monacoClosedPath = closedPath; monacoLivePath = livePath;
            try
            {
                Assert.AreEqual(1, Convert.ToInt32(((dynamic)workbooks).Count));
                ((dynamic)workbook).SaveAs(closedPath, 52);
                monacoLiveWorkbook = ((dynamic)workbooks).Add();
                ((dynamic)monacoLiveWorkbook).SaveAs(livePath, 52);
                RequireMonacoBook(workbook, closedPath); RequireMonacoBook(monacoLiveWorkbook, livePath);
                RequireOnlyRegisteredMonacoBooks();
            }
            catch { PreserveMonacoNativeOutcome(); throw; }
        }

        internal void SaveMonacoBaselines(string closedPath, string livePath)
        {
            RequireMonacoBook(workbook, closedPath); RequireMonacoBook(monacoLiveWorkbook, livePath);
            RequireOnlyRegisteredMonacoBooks();
            try { ((dynamic)workbook).Save(); ((dynamic)monacoLiveWorkbook).Save(); }
            catch { PreserveMonacoNativeOutcome(); throw; }
        }

        internal void CloseMonacoOldScopeWithoutSave(string path)
        {
            RequireMonacoBook(workbook, path); RequireOnlyRegisteredMonacoBooks();
            Assert.IsFalse(monacoPrimaryCloseInvoked, "The old workbook close must not be replayed.");
            monacoPrimaryCloseInvoked = true;
            try { ((dynamic)workbook).Close(false); }
            catch { PreserveMonacoNativeOutcome(); throw; }
            Release(workbook); workbook = null;
            Assert.AreEqual(1, Convert.ToInt32(((dynamic)workbooks).Count));
        }

        internal void CloseMonacoLiveScopeWithoutSave(string path)
        {
            if (monacoLiveWorkbook == null) return;
            RequireMonacoBook(monacoLiveWorkbook, path); RequireOnlyRegisteredMonacoBooks();
            Assert.IsFalse(monacoLiveCloseInvoked, "The live workbook close must not be replayed.");
            monacoLiveCloseInvoked = true;
            try { ((dynamic)monacoLiveWorkbook).Close(false); }
            catch { PreserveMonacoNativeOutcome(); throw; }
            Release(monacoLiveWorkbook); monacoLiveWorkbook = null;
            RequireOnlyRegisteredMonacoBooks();
        }

        private void RequireOnlyRegisteredMonacoBooks()
        {
            RequireMonacoOwner();
            var expected = new List<string>();
            if (workbook != null) expected.Add(monacoClosedPath);
            if (monacoLiveWorkbook != null) expected.Add(monacoLivePath);
            Assert.AreEqual(expected.Count, Convert.ToInt32(((dynamic)workbooks).Count), "An unregistered workbook appeared; native cleanup must refuse.");
            for (int index = 1; index <= expected.Count; index++)
            {
                object book = ((dynamic)workbooks).Item(index);
                try
                {
                    string fullName = Convert.ToString(((dynamic)book).FullName);
                    Assert.IsTrue(expected.Contains(fullName, StringComparer.OrdinalIgnoreCase));
                }
                finally { BalanceMonacoGetter(book); }
            }
        }

        /// <summary>Matches UiText's catalogue selection using owned native VBE menu captions, with balanced getter lifetimes.</summary>
        internal CultureInfo ReadMonacoHostCulture()
        {
            RequireMonacoOwner(); object editor = null, bars = null;
            var captions = new List<string>(); bool menu = false;
            try
            {
                editor = ((dynamic)application).VBE; bars = ((dynamic)editor).CommandBars;
                int count = Convert.ToInt32(((dynamic)bars).Count); Assert.IsTrue(count >= 0 && count <= 1000);
                for (int index = 1; index <= count; index++)
                {
                    object bar = null, controls = null;
                    try
                    {
                        bar = ((dynamic)bars)[index]; if (Convert.ToInt32(((dynamic)bar).Type) != 1) continue;
                        menu = true; controls = ((dynamic)bar).Controls;
                        int controlCount = Convert.ToInt32(((dynamic)controls).Count); Assert.IsTrue(controlCount >= 0 && controlCount <= 1000);
                        for (int controlIndex = 1; controlIndex <= controlCount; controlIndex++)
                        {
                            object control = ((dynamic)controls)[controlIndex];
                            try { captions.Add(Convert.ToString(((dynamic)control).Caption)); }
                            finally { BalanceMonacoGetter(control); }
                        }
                    }
                    finally { BalanceMonacoGetter(controls); BalanceMonacoGetter(bar); }
                }
                return menu ? UiLanguages.FromMenus(captions, CultureInfo.CurrentUICulture) : UiText.Supported(CultureInfo.CurrentUICulture);
            }
            finally { BalanceMonacoGetter(bars); BalanceMonacoGetter(editor); }
        }

        internal static string MonacoLocalized(string text, CultureInfo culture)
        {
            var language = UiLanguages.All.Single(l => l.CultureName == UiText.Supported(culture).Name);
            var catalogue = new ResourceManager("VBAi.Localization.UiStrings" + language.ResourceSuffix, typeof(UiText).Assembly);
            try { return catalogue.GetString(text, CultureInfo.InvariantCulture) ?? text; }
            finally { catalogue.ReleaseAllResources(); }
        }

        internal MonacoNativeObservation ReadInstalledMonaco(CultureInfo culture)
        {
            RequireMonacoOwner(); object editor = null, main = null;
            var result = new MonacoNativeObservation();
            try
            {
                editor = ((dynamic)application).VBE; main = ((dynamic)editor).MainWindow;
                IntPtr vbe = new IntPtr(Convert.ToInt64(((dynamic)main).HWnd));
                uint pid; GetWindowThreadProcessId(vbe, out pid); result.VbeHandle = vbe.ToInt64(); result.VbeProcessId = (int)pid;
                Assert.AreEqual(ProcessId, result.VbeProcessId);
                IntPtr candidate = IntPtr.Zero; int candidates = 0;
                string caption = MonacoLocalized("VBAi editor", culture);
                EnumChildWindows(vbe, (window, unused) =>
                {
                    var title = new StringBuilder(512); var className = new StringBuilder(256);
                    GetWindowText(window, title, title.Capacity); GetClassName(window, className, className.Capacity);
                    if (title.ToString() != caption || !className.ToString().StartsWith("WindowsForms", StringComparison.Ordinal)) return true;
                    GetWindowThreadProcessId(window, out pid);
                    if (pid != (uint)ProcessId) { result.EditorProcessId = (int)pid; candidates = 2; return false; }
                    candidate = window; candidates++; return true;
                }, IntPtr.Zero);
                Assert.IsTrue(candidates <= 1, "The exact owned VBE contains ambiguous editor windows.");
                if (candidate == IntPtr.Zero) return result;
                result.EditorHandle = candidate.ToInt64(); result.EditorProcessId = ProcessId;
                result.Embedded = IsChild(vbe, candidate); result.Visible = IsWindowVisible(candidate);
                var nativeClass = new StringBuilder(256); GetClassName(candidate, nativeClass, nativeClass.Capacity);
                result.EditorClass = nativeClass.ToString(); result.EditorCaption = caption;
                var root = AutomationElement.FromHandle(candidate);
                Assert.AreEqual(ProcessId, root.Current.ProcessId);
                var pending = new Queue<AutomationElement>(); pending.Enqueue(root);
                var names = new List<string>(); var selected = new List<string>(); var tabPids = new List<int>();
                var textPids = new HashSet<int>(); var text = new StringBuilder();
                var walker = TreeWalker.ControlViewWalker;
                while (pending.Count != 0 && result.ObservedElements < 512)
                {
                    var current = pending.Dequeue(); result.ObservedElements++;
                    var state = current.Current;
                    if (state.ProcessId == ProcessId && (state.AutomationId == "status" ||
                        (state.ControlType == ControlType.Text && (state.ClassName ?? "").StartsWith("WindowsForms10.STATIC", StringComparison.Ordinal))))
                    {
                        result.StatusCount++; result.StatusText = state.Name; result.StatusProcessId = state.ProcessId;
                        result.StatusAutomationId = state.AutomationId; result.StatusVisible = !state.IsOffscreen;
                        result.StatusClassName = state.ClassName;
                    }
                    if (state.ControlType == ControlType.TabItem && state.ProcessId == ProcessId)
                    {
                        names.Add(state.Name); tabPids.Add(state.ProcessId);
                        object pattern;
                        if (current.TryGetCurrentPattern(SelectionItemPattern.Pattern, out pattern) && ((SelectionItemPattern)pattern).Current.IsSelected)
                            selected.Add(state.Name);
                    }
                    if (!state.IsOffscreen && !state.IsPassword && text.Length < 32768)
                    {
                        // Browser providers may run in a WebView child process. Read only this editor's
                        // descendant tree; their PID is evidence, never ownership or an action target.
                        AppendMonacoText(text, state.Name); textPids.Add(state.ProcessId);
                        object pattern;
                        if (text.Length < 32768 && current.TryGetCurrentPattern(TextPattern.Pattern, out pattern))
                            AppendMonacoText(text, ((TextPattern)pattern).DocumentRange.GetText(Math.Min(8192, 32768 - text.Length)));
                        else if (text.Length < 32768 && current.TryGetCurrentPattern(ValuePattern.Pattern, out pattern))
                            AppendMonacoText(text, ((ValuePattern)pattern).Current.Value);
                    }
                    for (var child = walker.GetFirstChild(current); child != null; child = walker.GetNextSibling(child))
                    {
                        if (pending.Count >= 512) { result.TreeTruncated = true; break; }
                        pending.Enqueue(child);
                    }
                }
                result.TreeTruncated |= pending.Count != 0; result.Tabs = names.ToArray(); result.SelectedTabs = selected.ToArray();
                result.TabProcessIds = tabPids.ToArray(); result.RenderedText = text.ToString(); result.TextProviderProcessIds = textPids.ToArray();
                return result;
            }
            catch (ElementNotAvailableException error) { result.ObservationError = error.ToString(); return result; }
            finally { BalanceMonacoGetter(main); BalanceMonacoGetter(editor); }
        }

        internal static string MonacoTabCaption(string project, string module)
        {
            Assert.IsFalse(string.IsNullOrWhiteSpace(project)); Assert.IsFalse(string.IsNullOrWhiteSpace(module));
            return project + " · " + module; // EditorVbeModule.Name, not a bare VBComponent name.
        }

        internal static string MonacoAccessibleSourceStatus(MonacoNativeObservation state)
        {
            string text = state?.RenderedText ?? "";
            return new[] { "Option Explicit", "LiveValue", "42", "VBAi owned live scope probe" }.All(text.Contains)
                ? "LIVE_SOURCE_MARKERS_OBSERVED" : "LIVE_SOURCE_NOT_EXPOSED_BY_ACCESSIBILITY";
        }

        /// <summary>Qualifies native status/tab ownership only; code rendering remains a separate genuine-image review.</summary>
        internal static void RequireInstalledMonacoObservation(MonacoNativeObservation state, int pid, string closedTab, string liveTab, string synchronized)
        {
            Assert.IsNotNull(state); Assert.IsTrue(pid > 0);
            Assert.IsTrue(state.VbeHandle != 0 && state.EditorHandle != 0 && state.VbeHandle != state.EditorHandle);
            Assert.AreEqual(pid, state.VbeProcessId); Assert.AreEqual(pid, state.EditorProcessId);
            Assert.IsTrue((state.EditorClass ?? "").StartsWith("WindowsForms", StringComparison.Ordinal));
            Assert.IsTrue(state.Embedded && state.Visible, "Only the visible actual editor embedded in the owned VBE qualifies.");
            Assert.IsNull(state.ObservationError); Assert.IsFalse(state.TreeTruncated);
            Assert.AreEqual(1, state.StatusCount);
            Assert.IsTrue(state.StatusAutomationId == "status" || (state.StatusClassName ?? "").StartsWith("WindowsForms10.STATIC", StringComparison.Ordinal));
            Assert.AreEqual(pid, state.StatusProcessId);
            Assert.IsTrue(state.StatusVisible);
            Assert.AreEqual(synchronized, state.StatusText, "Inactive closed-project error contaminated the live status, or live synchronization did not settle.");
            Assert.IsTrue(state.Tabs.Count(name => name == closedTab) == 1 && state.Tabs.Count(name => name == liveTab) == 1,
                "The exact project/module captions must be unique, with the closed draft retained.");
            Assert.IsTrue(state.TabProcessIds.Length == state.Tabs.Length && state.TabProcessIds.All(owner => owner == pid));
            CollectionAssert.AreEqual(new[] { liveTab }, state.SelectedTabs);
        }

        internal void CaptureInstalledMonaco(MonacoNativeObservation observed, string path)
        {
            var evidence = new Dictionary<string, object>
            {
                ["Scope"] = "Actual installed embedded editor; no detached ModernEditorWindow",
                ["ProcessId"] = ProcessId,
                ["AssemblyMvid"] = typeof(VbeSession).Module.ModuleVersionId.ToString("D"),
                ["Observation"] = observed,
                ["State"] = "PENDING",
                ["Path"] = path,
                ["Method"] = "PrintWindow, flags=2",
                ["VisualReview"] = "NOT_RUN",
                ["AccessibilitySource"] = MonacoAccessibleSourceStatus(observed),
                ["RenderedCodeVerified"] = false
            };
            try
            {
                RequireMonacoOwner(); IntPtr handle = new IntPtr(observed.EditorHandle);
                uint owner; GetWindowThreadProcessId(handle, out owner); Assert.AreEqual((uint)ProcessId, owner);
                Assert.IsTrue(IsChild(new IntPtr(observed.VbeHandle), handle) && IsWindowVisible(handle));
                GitCaptureRect bounds; Assert.IsTrue(GetWindowRect(handle, out bounds));
                int width = bounds.Right - bounds.Left, height = bounds.Bottom - bounds.Top;
                evidence["Width"] = width; evidence["Height"] = height;
                Assert.IsTrue(width > 0 && height > 0 && width <= 4096 && height <= 4096);
                using (var bitmap = new Bitmap(width, height))
                using (var graphics = Graphics.FromImage(bitmap))
                {
                    IntPtr dc = graphics.GetHdc();
                    try { Assert.IsTrue(PrintWindow(handle, dc, 2), "Native embedded editor capture failed."); }
                    finally { graphics.ReleaseHdc(dc); }
                    bitmap.Save(path, ImageFormat.Png);
                }
                GetWindowThreadProcessId(handle, out owner); Assert.AreEqual((uint)ProcessId, owner);
                Assert.IsTrue(IsChild(new IntPtr(observed.VbeHandle), handle));
                using (var sha = System.Security.Cryptography.SHA256.Create())
                using (var file = System.IO.File.OpenRead(path)) evidence["CaptureSha256"] = BitConverter.ToString(sha.ComputeHash(file)).Replace("-", "");
                evidence["State"] = "CAPTURED_PENDING_VISUAL_REVIEW";
            }
            catch (Exception error) { evidence["State"] = "FAILED"; evidence["Error"] = error.ToString(); throw; }
            finally { System.IO.File.WriteAllText(path + ".json", new System.Web.Script.Serialization.JavaScriptSerializer().Serialize(evidence), new UTF8Encoding(false)); }
        }

        private static void BalanceMonacoGetter(object value)
        { if (value != null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value); }

        private static void AppendMonacoText(StringBuilder output, string value)
        {
            if (string.IsNullOrEmpty(value) || output.Length >= 32768) return;
            output.Append(value, 0, Math.Min(value.Length, 32768 - output.Length));
            if (output.Length < 32768) output.Append('\n');
        }
    }
}
