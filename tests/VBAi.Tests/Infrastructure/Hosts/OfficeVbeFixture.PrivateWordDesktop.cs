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
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Integration
{
    internal sealed partial class OfficeVbeFixture
    {
        private const uint NativeObjectModel = 0xFFFFFFF0;
        [DllImport("user32.dll")] private static extern bool EnumChildWindows(IntPtr parent, DialogWindowCallback callback, IntPtr state);
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
            // Microsoft documents /n as a fresh Word instance, /q as no splash, and /m as no AutoExec.
            // The only file argument is the freshly created macro-free DOCX in this owned results directory.
            privateWordChild = IsolatedTestDesktop.Launch(executable, WordPrivateArguments(seed), Root, desktopName);
            ProcessId = privateWordChild.ProcessId;
            var wordProcesses = Process.GetProcessesByName("WINWORD");
            try
            {
                Assert.AreEqual(1, wordProcesses.Length, "Private Word ownership requires one process after the owned launch.");
                Assert.AreEqual(ProcessId, wordProcesses[0].Id);
            }
            finally { foreach (var process in wordProcesses) process.Dispose(); }
            Assert.AreEqual(desktopName, IsolatedTestDesktop.DesktopName(privateWordChild.ThreadId), true,
                "The launched Word primary thread is outside the exact private desktop.");
            CaptureOwnedProcess();
            Assert.AreEqual(ProcessId, ownedProcess.Id);
            Assert.AreEqual(executable, ExcelOwnedProcessImage.Read(ownedProcess.Handle), true,
                "The original launched Word process image differs from the reviewed executable.");
            Assert.IsFalse(privateWordChild.Wait(0), "The original private Word process exited before NativeOM attachment.");
            AttachOnlyLaunchedWordNativeObjectModel(seed, desktopName);
            owned = true;
            privateWordSeed = document;
            steps.Add(new { ApplicationOwnershipVerifiedBeforeMutation = true, Method = "LaunchedPrivateWordNativeObjectModel",
                ProcessId, LaunchThreadId = privateWordChild.ThreadId, Desktop = desktopName,
                SeedPath = seed, SeedSha256 = HashOwnedFile(seed), Executable = executable,
                ExecutableSha256 = expectedHash.ToUpperInvariant(), ForceTermination = false });
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
            return new[] { "/n", "/q", "/m", Path.GetFullPath(seed) };
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

        private IntPtr[] LaunchedWordDocumentWindows()
        {
            var found = new HashSet<IntPtr>(); int topVisited = 0, childVisited = 0;
            DialogWindowCallback top = (window, state) => {
                if (++topVisited > 4096) return false;
                uint pid; GetWindowThreadProcessId(window, out pid);
                if (pid != (uint)ProcessId) return true;
                EnumChildWindows(window, (child, unused) => {
                    if (++childVisited > 2048) return false;
                    uint childPid; GetWindowThreadProcessId(child, out childPid);
                    var className = new StringBuilder(64); GetClassName(child, className, className.Capacity);
                    if (childPid == (uint)ProcessId && className.ToString() == "_WwG") found.Add(child);
                    return true;
                }, IntPtr.Zero);
                return childVisited <= 2048;
            };
            bool completed = EnumWindows(top, IntPtr.Zero);
            if (!completed || topVisited > 4096 || childVisited > 2048 || found.Count > 8)
                throw new InvalidOperationException("Bounded private Word NativeOM window inventory failed.");
            return found.ToArray();
        }

        private void AttachOnlyLaunchedWordNativeObjectModel(string seedPath, string desktopName)
        {
            var clock = Stopwatch.StartNew(); IntPtr documentWindow = IntPtr.Zero;
            while (clock.Elapsed < TimeSpan.FromSeconds(30))
            {
                if (privateWordChild.Wait(0)) throw new InvalidOperationException("Original private Word exited before NativeOM attachment.");
                IntPtr[] windows = LaunchedWordDocumentWindows();
                if (windows.Length > 1) throw new InvalidOperationException("Ambiguous launched-PID Word NativeOM document windows.");
                if (windows.Length == 1) { documentWindow = windows[0]; break; }
                Thread.Sleep(100);
            }
            if (documentWindow == IntPtr.Zero) throw new TimeoutException("No unique launched-PID Word NativeOM window appeared.");
            uint pid; uint uiThread = GetWindowThreadProcessId(documentWindow, out pid);
            Assert.AreEqual((uint)ProcessId, pid);
            Assert.AreEqual(desktopName, IsolatedTestDesktop.DesktopName(uiThread), true,
                "The actual Word document UI thread is outside the exact private desktop.");
            object native = null, app = null, documents = null, seed = null, active = null, window = null;
            IntPtr seedUnknown = IntPtr.Zero, activeUnknown = IntPtr.Zero;
            try
            {
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
    }
}
