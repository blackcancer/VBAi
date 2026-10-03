using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Web.Script.Serialization;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Integration
{
    internal sealed partial class OfficeVbeFixture
    {
        private const uint NativeObjectModel = 0xFFFFFFF0;
        [DllImport("user32.dll")] private static extern bool EnumChildWindows(IntPtr parent, DialogWindowCallback callback, IntPtr state);
        [DllImport("user32.dll")] private static extern IntPtr GetAncestor(IntPtr window, uint flags);
        [DllImport("oleacc.dll", EntryPoint = "AccessibleObjectFromWindow")]
        private static extern int NativeAccessibleObjectFromWindow(IntPtr window, uint objectId, ref Guid iid,
            [MarshalAs(UnmanagedType.Interface)] out object value);

        private IsolatedTestDesktop.NativeChild privateWordChild;
        private object privateWordSeed;

        internal static string RequireEnabledOfficeDesktop(string kind, string optIn, string required,
            string configured, Action<string> requireCurrent)
        {
            if (!string.Equals(optIn, "1", StringComparison.Ordinal))
                Assert.Inconclusive("Set VBAi_RUN_OFFICE_TESTS=1 to qualify installed Office hosts.");
            return RequirePrivateWordDesktop(kind, required, configured, requireCurrent);
        }

        internal static string RequirePrivateWordDesktop(string kind, string required, string configured,
            Action<string> requireCurrent)
        {
            if (requireCurrent == null) throw new ArgumentNullException(nameof(requireCurrent));
            if (!string.IsNullOrEmpty(required) && !string.Equals(configured, required, StringComparison.Ordinal))
                throw new InvalidOperationException("The qualification desktop and Word test desktop must match exactly.");
            string selected = !string.IsNullOrEmpty(required) ? required : configured;
            if (string.IsNullOrEmpty(selected)) return null;
            if (!string.Equals(kind, "Word", StringComparison.Ordinal))
                throw new InvalidOperationException("The private-desktop bootstrap is prepared only for owned Word qualification.");
            requireCurrent(selected);
            return selected;
        }

        private void BootstrapPrivateWordDesktop(string desktopName)
        {
            IsolatedTestDesktop.RequireCurrent(desktopName);
            string expectedHash = Environment.GetEnvironmentVariable("VBAi_TEST_WORD_EXE_SHA256");
            string executable = RequireExactWordExecutable(Environment.GetEnvironmentVariable("VBAi_TEST_WORD_EXE"),
                expectedHash, File.Exists, HashOwnedFile);
            string seed = Path.Combine(Root, "NativeObjectModelSeed.docx");
            WriteMacroFreeSeed(seed);
            // /a prevents automatic loading of Normal/global templates before NativeOM attachment.
            // Use one documented switch; splash windows remain on the inactive desktop.
            // The only file argument is the freshly created macro-free DOCX in this owned results directory.
            privateWordChild = IsolatedTestDesktop.Launch(executable, WordPrivateArguments(seed), Root, desktopName);
            ProcessId = privateWordChild.ProcessId;
            PreparePrivateWordLaunch(desktopName, privateWordChild.ThreadId,
                () => PersistPrivateWordLaunch(seed, executable, expectedHash, desktopName),
                () => {
                    Assert.IsFalse(privateWordChild.Wait(0), "Original Word exited before its identity capture.");
                    CaptureOwnedProcess();
                    Assert.AreEqual(ProcessId, ownedProcess.Id);
                    Assert.AreEqual(executable, ExcelOwnedProcessImage.Read(privateWordChild.ProcessHandle), true);
                    Assert.IsFalse(privateWordChild.Wait(0), "Original Word exited during its identity capture.");
                    steps.Add(new { Phase = "PrivateWordOriginalProcessCaptured", Identity = shutdownEvidence.Record });
                    FlushAdapterEvidence();
                },
                () => {
                    var wordProcesses = Process.GetProcessesByName("WINWORD");
                    try
                    {
                        Assert.AreEqual(1, wordProcesses.Length, "Private Word ownership requires one process after the owned launch.");
                        Assert.AreEqual(ProcessId, wordProcesses[0].Id);
                    }
                    finally { foreach (var process in wordProcesses) process.Dispose(); }
                },
                () => IsolatedTestDesktop.ReadThreadDesktop(privateWordChild.ThreadId),
                observation => {
                    steps.Add(new { Phase = "PrivateWordPrimaryThreadObservation", Observation = observation,
                        PlacementProved = false, ActualDocumentUiDesktopRequired = true });
                    FlushAdapterEvidence();
                },
                () => AttachOnlyLaunchedWordNativeObjectModel(seed, desktopName));
            owned = true;
            privateWordSeed = document;
            steps.Add(new { ApplicationOwnershipVerifiedBeforeMutation = true, Method = "LaunchedPrivateWordNativeObjectModel",
                ProcessId, LaunchThreadId = privateWordChild.ThreadId, Desktop = desktopName,
                SeedPath = seed, SeedSha256 = HashOwnedFile(seed), Executable = executable,
                ExecutableSha256 = expectedHash.ToUpperInvariant(), ForceTermination = false });
        }

        private void PersistPrivateWordLaunch(string seed, string executable, string hash, string desktop)
        {
            var receipt = new { Phase = "PrivateWordCreateProcessReturned", ProcessId,
                LaunchThreadId = privateWordChild.ThreadId,
                OriginalLaunchProcessHandle = privateWordChild.ProcessHandle.ToInt64(),
                Desktop = desktop, Executable = executable, ExecutableSha256 = hash,
                RequestedCommandLine = IsolatedTestDesktop.CommandLine(executable, WordPrivateArguments(seed)),
                SeedPath = seed, SeedSha256 = HashOwnedFile(seed),
                OwnershipVerified = false, ComCalls = 0, UiActions = 0,
                ExpectedAssemblyMvid = typeof(VbeSession).Module.ModuleVersionId.ToString("D"),
                Utc = DateTime.UtcNow.ToString("o") };
            using (var stream = new FileStream(Path.Combine(Root, "private-word-launch.json"), FileMode.CreateNew, FileAccess.Write))
            using (var writer = new StreamWriter(stream, new UTF8Encoding(false)))
                writer.Write(new JavaScriptSerializer().Serialize(receipt));
            steps.Add(receipt);
        }

        /// <summary>Persists launch and original process identity before discovery; only the actual document UI thread can prove placement.</summary>
        internal static void PreparePrivateWordLaunch(string desktop, uint primaryThread,
            Action persistLaunch, Action captureIdentity, Action inventory,
            Func<IsolatedTestDesktop.ThreadDesktopObservation> readPrimary,
            Action<IsolatedTestDesktop.ThreadDesktopObservation> persistPrimary, Action attach)
        {
            if (persistLaunch == null || captureIdentity == null || inventory == null || readPrimary == null ||
                persistPrimary == null || attach == null) throw new ArgumentNullException("Private Word launch dependencies");
            IsolatedTestDesktop.RequireName(desktop);
            if (primaryThread == 0) throw new ArgumentException("The original launched thread must be recorded.");
            persistLaunch();
            captureIdentity();
            inventory();
            var observation = readPrimary();
            if (observation == null || observation.ThreadId != primaryThread)
                throw new InvalidOperationException("The launch-thread observation belongs to a different generation.");
            persistPrimary(observation);
            if (!string.IsNullOrEmpty(observation.Name) && (observation.Handle == 0 ||
                !string.Equals(observation.Name, desktop, StringComparison.OrdinalIgnoreCase)))
                throw new InvalidOperationException("The observed primary Word thread belongs to another desktop.");
            attach(); // Requires exact HWND/PID/actual UI thread desktop BEFORE the first COM call.
        }

        internal static string RequireExactWordExecutable(string path, string expectedHash,
            Func<string, bool> exists, Func<string, string> actualHash)
        {
            if (exists == null) throw new ArgumentNullException(nameof(exists));
            if (actualHash == null) throw new ArgumentNullException(nameof(actualHash));
            if (string.IsNullOrWhiteSpace(path) || !Path.IsPathRooted(path) ||
                !string.Equals(Path.GetFileName(path), "WINWORD.EXE", StringComparison.OrdinalIgnoreCase) ||
                expectedHash == null || expectedHash.Length != 64 || expectedHash.Any(value => !Uri.IsHexDigit(value)))
                throw new InvalidOperationException("An exact installed WINWORD.EXE path and SHA-256 are required for private-desktop qualification.");
            string canonical = Path.GetFullPath(path);
            if (!exists(canonical) || !string.Equals(actualHash(canonical), expectedHash, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The installed Word executable changed before private-desktop launch.");
            return canonical;
        }

        internal static string[] WordPrivateArguments(string seed)
        {
            if (string.IsNullOrWhiteSpace(seed) || !Path.IsPathRooted(seed) ||
                !string.Equals(Path.GetExtension(seed), ".docx", StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("Only an absolute macro-free DOCX seed may be opened on the private Word desktop.");
            return new[] { "/a", Path.GetFullPath(seed) };
        }

        private static void WriteMacroFreeSeed(string path)
        {
            Assert.IsFalse(File.Exists(path));
            const string contentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
            const string relationship = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
            const string xml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body><w:p/></w:body></w:document>";
            using (var package = Package.Open(path, FileMode.CreateNew, FileAccess.ReadWrite))
            {
                var part = package.CreatePart(new Uri("/word/document.xml", UriKind.Relative), contentType);
                byte[] bytes = Encoding.UTF8.GetBytes(xml);
                using (var stream = part.GetStream(FileMode.Create, FileAccess.Write)) stream.Write(bytes, 0, bytes.Length);
                package.CreateRelationship(part.Uri, TargetMode.Internal, relationship);
            }
            using (var package = Package.Open(path, FileMode.Open, FileAccess.Read))
            {
                var contentParts = package.GetParts().Where(part => !PackUriHelper.IsRelationshipPartUri(part.Uri)).ToArray();
                Assert.AreEqual(1, contentParts.Length, "The seed DOCX must have exactly one content part and no macros.");
                Assert.AreEqual(new Uri("/word/document.xml", UriKind.Relative), contentParts[0].Uri);
                Assert.AreEqual(contentType, contentParts[0].ContentType);
                var relationships = package.GetRelationships().ToArray();
                Assert.AreEqual(1, relationships.Length);
                Assert.AreEqual(relationship, relationships[0].RelationshipType);
                Assert.AreEqual(TargetMode.Internal, relationships[0].TargetMode);
                Assert.AreEqual(contentParts[0].Uri, relationships[0].TargetUri);
            }
        }

        private static string HashOwnedFile(string path)
        {
            using (var file = File.OpenRead(path)) using (var sha = SHA256.Create())
                return BitConverter.ToString(sha.ComputeHash(file)).Replace("-", "");
        }

        private IntPtr[] LaunchedWordDocumentWindows(string desktop)
        {
            return CollectPrivateWordDocumentWindows((uint)ProcessId,
                callback => IsolatedTestDesktop.InventoryWindows(desktop, callback),
                (root, callback) => {
                    Exception failure = null;
                    EnumChildWindows(root, (child, unused) => {
                        try { return callback(child); }
                        catch (Exception error) { failure = error; return false; }
                    }, IntPtr.Zero);
                    if (failure != null) System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(failure).Throw();
                },
                window => { uint pid; GetWindowThreadProcessId(window, out pid); return pid; },
                window => { var name = new StringBuilder(64); GetClassName(window, name, name.Capacity); return name.ToString(); });
        }

        internal static IntPtr[] CollectPrivateWordDocumentWindows(uint expectedPid, Action<Func<IntPtr, bool>> visitRoots,
            Action<IntPtr, Func<IntPtr, bool>> visitChildren, Func<IntPtr, uint> readPid, Func<IntPtr, string> readClass)
        {
            if (expectedPid == 0) throw new ArgumentException("An exact launched Word PID is required.");
            if (visitRoots == null || visitChildren == null || readPid == null || readClass == null)
                throw new ArgumentNullException("Private Word window inventory dependencies");
            var found = new HashSet<IntPtr>(); int topVisited = 0, childVisited = 0;
            visitRoots(window => {
                if (++topVisited > 4096) return false;
                if (readPid(window) != expectedPid) return true;
                visitChildren(window, child => {
                    if (++childVisited > 2048) return false;
                    if (readPid(child) == expectedPid && readClass(child) == "_WwG") found.Add(child);
                    return true;
                });
                return childVisited <= 2048;
            });
            if (topVisited > 4096 || childVisited > 2048 || found.Count > 8)
                throw new InvalidOperationException("Bounded private Word NativeOM window inventory failed.");
            return found.ToArray();
        }

        private void AttachOnlyLaunchedWordNativeObjectModel(string seedPath, string desktopName)
        {
            var clock = Stopwatch.StartNew();
            IntPtr documentWindow = WaitForPrivateWordWindow(() => privateWordChild.Wait(0),
                () => LaunchedWordDocumentWindows(desktopName), () => clock.Elapsed < TimeSpan.FromSeconds(30),
                () => Thread.Sleep(100), documentHandle => RequirePrivateWordDocumentWindow(documentHandle, desktopName),
                documentHandle => ObservePrivateWordWindowReadiness(documentHandle, desktopName));
            uint pid; uint uiThread = GetWindowThreadProcessId(documentWindow, out pid);
            Assert.AreEqual((uint)ProcessId, pid);
            Assert.AreEqual(desktopName, IsolatedTestDesktop.DesktopName(uiThread), true,
                "The actual Word document UI thread is outside the exact private desktop.");
            object native = null, app = null, documents = null, seed = null, active = null, window = null;
            IntPtr seedUnknown = IntPtr.Zero, activeUnknown = IntPtr.Zero;
            try
            {
                IsolatedTestDesktop.RequireCurrent(desktopName);
                RequirePrivateWordDocumentWindow(documentWindow, desktopName);
                Assert.IsFalse(privateWordChild.Wait(0), "Original Word exited before the NativeOM call.");
                Guid dispatch = new Guid("00020400-0000-0000-C000-000000000046");
                int hr = NativeAccessibleObjectFromWindow(documentWindow, NativeObjectModel, ref dispatch, out native);
                if (hr != 0 || native == null) throw new InvalidOperationException("Launched-PID Word NativeOM attachment refused with HRESULT 0x" +
                    unchecked((uint)hr).ToString("X8") + ".");
                app = ((dynamic)native).Application;
                var probe = new VbeProjectComponents.NativeOtherHostProbe { ReadHostKind = () => "Word" };
                Assert.AreEqual((uint)ProcessId, probe.ApplicationProcessId(app),
                    "NativeOM returned an application outside the launched Word PID.");
                documents = ((dynamic)app).Documents;
                Assert.AreEqual(1, Convert.ToInt32(((dynamic)documents).Count), "Only the inert owned seed document may be open before mutation.");
                seed = ((dynamic)documents).Item(1);
                Assert.AreEqual(seedPath, Path.GetFullPath((string)((dynamic)seed).FullName), true);
                active = ((dynamic)app).ActiveDocument;
                seedUnknown = Marshal.GetIUnknownForObject(seed); activeUnknown = Marshal.GetIUnknownForObject(active);
                Assert.AreEqual(seedUnknown, activeUnknown);
                window = ((dynamic)app).ActiveWindow;
                IntPtr activeHandle = new IntPtr(Convert.ToInt64(((dynamic)window).hWnd));
                uint activePid; uint activeThread = GetWindowThreadProcessId(activeHandle, out activePid);
                Assert.AreEqual((uint)ProcessId, activePid);
                Assert.AreEqual(desktopName, IsolatedTestDesktop.DesktopName(activeThread), true);
                application = app; document = seed;
                app = null; seed = null;
                steps.Add(new { Phase = "PrivateWordNativeObjectModelAttached", ProcessId,
                    NativeObjectModelHandle = documentWindow.ToInt64(), UiThreadId = uiThread,
                    ActiveWindowHandle = activeHandle.ToInt64(), ActiveWindowThreadId = activeThread,
                    Desktop = desktopName, SeedPath = seedPath, OpenDocumentCount = 1 });
            }
            finally
            {
                if (activeUnknown != IntPtr.Zero) Marshal.Release(activeUnknown);
                if (seedUnknown != IntPtr.Zero) Marshal.Release(seedUnknown);
                Release(window);
                // ActiveDocument may alias the seed RCW that is handed to the fixture.
                if (active != null && Marshal.IsComObject(active)) Marshal.ReleaseComObject(active);
                Release(documents);
                if (native != null && Marshal.IsComObject(native)) Marshal.ReleaseComObject(native);
                Release(seed); Release(app);
            }
        }

        internal static IntPtr WaitForPrivateWordWindow(Func<bool> exited, Func<IntPtr[]> inventory,
            Func<bool> withinBound, Action wait, Action<IntPtr> verify, Func<IntPtr, bool> ready = null)
        {
            if (exited == null || inventory == null || withinBound == null || wait == null || verify == null)
                throw new ArgumentNullException("Private Word window observation dependencies");
            while (withinBound())
            {
                if (exited()) throw new InvalidOperationException("Original private Word exited before NativeOM attachment.");
                var windows = inventory();
                if (windows == null || windows.Length > 1)
                    throw new InvalidOperationException("Ambiguous or incomplete launched-PID Word NativeOM document windows.");
                if (windows.Length == 1)
                {
                    if (windows[0] == IntPtr.Zero) throw new InvalidOperationException("The Word document HWND is unavailable.");
                    if (ready != null && !ready(windows[0]))
                    {
                        wait(); // Only a positively owned, still-hidden startup window may be observed again.
                        continue;
                    }
                    verify(windows[0]);
                    if (exited()) throw new InvalidOperationException("Original Word exited during window verification.");
                    return windows[0];
                }
                wait(); // Observation only: no COM, activation or native action is retried.
            }
            throw new TimeoutException("No unique launched-PID Word NativeOM window appeared.");
        }

        private void RequirePrivateWordDocumentWindow(IntPtr handle, string desktop)
        {
            IntPtr root = GetAncestor(handle, 2);
            uint pid, rootPid;
            uint tid = GetWindowThreadProcessId(handle, out pid);
            uint rootTid = GetWindowThreadProcessId(root, out rootPid);
            var childClass = new StringBuilder(64); var rootClass = new StringBuilder(64);
            GetClassName(handle, childClass, childClass.Capacity);
            GetClassName(root, rootClass, rootClass.Capacity);
            RequirePrivateWordWindowIdentity(handle, root, pid, rootPid, tid, rootTid,
                childClass.ToString(), rootClass.ToString(), IsWindowVisible(handle), IsWindowVisible(root),
                (uint)ProcessId, desktop, IsolatedTestDesktop.DesktopName);
        }

        private bool ObservePrivateWordWindowReadiness(IntPtr handle, string desktop)
        {
            IntPtr root = GetAncestor(handle, 2);
            uint pid, rootPid;
            uint tid = GetWindowThreadProcessId(handle, out pid);
            uint rootTid = GetWindowThreadProcessId(root, out rootPid);
            var childClass = new StringBuilder(64); var rootClass = new StringBuilder(64);
            GetClassName(handle, childClass, childClass.Capacity);
            GetClassName(root, rootClass, rootClass.Capacity);
            bool visible = IsWindowVisible(handle), rootVisible = IsWindowVisible(root);
            steps.Add(new { Phase = "PrivateWordWindowReadinessObserved", Handle = handle.ToInt64(),
                Root = root.ToInt64(), ProcessId = pid, RootProcessId = rootPid, ThreadId = tid,
                RootThreadId = rootTid, ChildClass = childClass.ToString(), RootClass = rootClass.ToString(),
                Visible = visible, RootVisible = rootVisible, ComCalls = 0, UiActions = 0 });
            FlushAdapterEvidence();
            return PrivateWordWindowReady(handle, root, pid, rootPid, tid, rootTid,
                childClass.ToString(), rootClass.ToString(), visible, rootVisible,
                (uint)ProcessId, desktop, IsolatedTestDesktop.DesktopName);
        }

        internal static bool PrivateWordWindowReady(IntPtr handle, IntPtr root, uint pid, uint rootPid,
            uint tid, uint rootTid, string childClass, string rootClass, bool visible, bool rootVisible,
            uint expectedPid, string desktop, Func<uint, string> readDesktop)
        {
            if (readDesktop == null) throw new ArgumentNullException(nameof(readDesktop));
            IsolatedTestDesktop.RequireName(desktop);
            RequirePrivateWordWindowStructure(handle, root, pid, rootPid, tid, rootTid, childClass, rootClass, expectedPid);
            if (!string.Equals(readDesktop(tid), desktop, StringComparison.Ordinal))
                throw new InvalidOperationException("The actual Word UI thread is outside the exact private desktop.");
            return visible && rootVisible;
        }

        internal static void RequirePrivateWordWindowIdentity(IntPtr handle, IntPtr root, uint pid, uint rootPid,
            uint tid, uint rootTid, string childClass, string rootClass, bool visible, bool rootVisible,
            uint expectedPid, string desktop, Func<uint, string> readDesktop)
        {
            if (readDesktop == null) throw new ArgumentNullException(nameof(readDesktop));
            IsolatedTestDesktop.RequireName(desktop);
            RequirePrivateWordWindowStructure(handle, root, pid, rootPid, tid, rootTid, childClass, rootClass, expectedPid);
            if (!visible || !rootVisible)
                throw new InvalidOperationException("The Word document/root is not yet visible; NativeOM is not permitted.");
            if (!string.Equals(readDesktop(tid), desktop, StringComparison.Ordinal))
                throw new InvalidOperationException("The actual Word UI thread is outside the exact private desktop.");
        }

        private static void RequirePrivateWordWindowStructure(IntPtr handle, IntPtr root, uint pid, uint rootPid,
            uint tid, uint rootTid, string childClass, string rootClass, uint expectedPid)
        {
            if (expectedPid == 0 || handle == IntPtr.Zero || root == IntPtr.Zero || handle == root ||
                pid != expectedPid || rootPid != expectedPid || tid == 0 || rootTid != tid ||
                childClass != "_WwG" || rootClass != "OpusApp")
                throw new InvalidOperationException("The Word document window/root identity is incomplete or foreign.");
        }
    }
}
