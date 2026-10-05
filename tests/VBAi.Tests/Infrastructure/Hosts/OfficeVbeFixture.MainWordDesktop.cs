using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Web.Script.Serialization;

namespace VBAi.Tests.Integration
{
    internal sealed partial class OfficeVbeFixture
    {
        internal const string MainWordEnvironment = "VBAi_RUN_WORD_MAIN_DESKTOP_TESTS";
        private bool mainWordDesktop;
        private string mainWordExecutable, mainWordBirth;

        internal static bool SelectMainWordDesktop(string kind, string optIn, string required, string configured, Action requireMain)
        {
            if (optIn != "1") return false;
            if (kind != "Word" || !string.IsNullOrEmpty(required) || !string.IsNullOrEmpty(configured))
                throw new InvalidOperationException("Explicit Main Word requires Word and both private desktop variables empty.");
            if (requireMain == null) throw new ArgumentNullException(nameof(requireMain));
            requireMain();
            return true;
        }

        private void RequireWordBootstrapDesktop(string desktop)
        {
            if (mainWordDesktop) IsolatedTestDesktop.RequireMainCurrent();
            else IsolatedTestDesktop.RequireCurrent(desktop);
        }

        internal static void PrepareMainWordLaunch(Action persistLaunch, Action captureIdentity,
            Action inventory, Action requireMain, Action attach)
        {
            if (persistLaunch == null || captureIdentity == null || inventory == null || requireMain == null || attach == null)
                throw new ArgumentNullException("Main Word launch dependencies");
            // Even a placement change after CreateProcess leaves a durable original-launch receipt.
            persistLaunch(); captureIdentity(); requireMain(); inventory(); requireMain(); attach();
        }

        internal static void ReplaceWordSeedOnce(Action requirePlacement, Action createDocument, Action closeSeed,
            Action releaseSeed, Action clearSeed)
        {
            if (requirePlacement == null || createDocument == null || closeSeed == null || releaseSeed == null || clearSeed == null)
                throw new ArgumentNullException("Owned Word seed transition dependencies");
            requirePlacement(); createDocument(); requirePlacement(); closeSeed(); releaseSeed(); clearSeed();
        }

        private void RequireMainWordOriginal()
        {
            IsolatedTestDesktop.RequireMainCurrent();
            if (privateWordChild == null || ownedProcess == null)
                throw new InvalidOperationException("The original Main Word generation/image is unavailable; no action or fallback is allowed.");
            RequireMainWordGeneration(ProcessId, mainWordBirth, mainWordExecutable,
                privateWordChild.ProcessHandle, privateWordChild.ProcessId, privateWordChild.Wait(0),
                ownedProcess.Id, ownedProcess.HasExited, ownedProcess.StartTime.ToUniversalTime().ToString("o"),
                ExcelOwnedProcessImage.Read(privateWordChild.ProcessHandle));
        }

        internal static void RequireMainWordGeneration(int expectedPid, string birth, string image, IntPtr originalHandle,
            int childPid, bool nativeExited, int capturedPid, bool capturedExited, string actualBirth, string actualImage)
        {
            if (expectedPid <= 0 || string.IsNullOrEmpty(birth) || string.IsNullOrEmpty(image) || originalHandle == IntPtr.Zero ||
                childPid != expectedPid || capturedPid != expectedPid || nativeExited || capturedExited || birth != actualBirth ||
                !string.Equals(image, actualImage, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The original Main Word generation/image is unavailable; no action or fallback is allowed.");
        }

        private void RequireMainWordBeforeNative()
        {
            if (mainWordDesktop) ObserveMainWord(IntPtr.Zero, true, false);
        }

        internal static IsolatedTestDesktop.MainInventory ObserveMainWordPlacement(Action requireOriginal,
            Func<IsolatedTestDesktop.MainInventory> read, uint processId, IntPtr requiredRoot, bool requireWord, bool requireVbe)
        {
            if (requireOriginal == null || read == null) throw new ArgumentNullException("Main placement dependencies");
            requireOriginal();
            var inventory = read();
            IsolatedTestDesktop.RequireMainInventory(inventory, processId, requiredRoot, requireWord, requireVbe);
            requireOriginal();
            return inventory;
        }

        private IsolatedTestDesktop.MainInventory ObserveMainWord(IntPtr requiredRoot, bool requireWord, bool requireVbe) =>
            ObserveMainWordPlacement(RequireMainWordOriginal, () => IsolatedTestDesktop.ReadMainWindows((uint)ProcessId),
                (uint)ProcessId, requiredRoot, requireWord, requireVbe);

        private void RequireMainWordWindow(IntPtr window, bool requireVisible, string expectedRootClass)
        {
            IntPtr root = GetAncestor(window, 2);
            uint pid, rootPid;
            uint tid = GetWindowThreadProcessId(window, out pid), rootTid = GetWindowThreadProcessId(root, out rootPid);
            if (window == IntPtr.Zero || root == IntPtr.Zero || pid != (uint)ProcessId || rootPid != pid || tid == 0 || rootTid != tid ||
                (requireVisible && (!IsWindowVisible(window) || !IsWindowVisible(root))))
                throw new InvalidOperationException("Main Word window/root identity or visibility is unavailable.");
            var inventory = ObserveMainWord(root, false, false);
            var observed = inventory.Windows.Single(row => row.Handle == root.ToInt64());
            RequireMainWordWindowObservation(observed, root, rootPid, rootTid, requireVisible, expectedRootClass);
        }

        internal static void RequireMainWordWindowObservation(IsolatedTestDesktop.MainWindow observed, IntPtr root,
            uint processId, uint threadId, bool requireVisible, string expectedRootClass)
        {
            if (observed == null || root == IntPtr.Zero || processId == 0 || threadId == 0 ||
                observed.Handle != root.ToInt64() || observed.ProcessId != processId || observed.ThreadId != threadId ||
                string.IsNullOrEmpty(expectedRootClass) || observed.ClassName != expectedRootClass || (requireVisible && !observed.Visible))
                throw new InvalidOperationException("The observed Default root identity changed.");
        }

        private void PersistMainWordReady(IDictionary<string, object> status)
        {
            if (!mainWordDesktop) return;
            string expectedMvid = Environment.GetEnvironmentVariable("VBAi_TEST_EMBEDDED_GIT_MVID");
            string expectedHash = Environment.GetEnvironmentVariable("VBAi_TEST_EMBEDDED_GIT_SHA256");
            string path = RequireMainWordProduct(status, ProcessId, expectedMvid, expectedHash,
                typeof(VbeSession).Module.ModuleVersionId, HashOwnedFile);
            var inventory = ObserveMainWord(IntPtr.Zero, true, true);
            var record = new { Phase = "MainWordReadyPlacementObserved", ProcessId, ProcessStartedUtc = mainWordBirth,
                Executable = mainWordExecutable, OriginalLaunchProcessHandle = privateWordChild.ProcessHandle.ToInt64(),
                OriginalProcessHandle = ownedProcess.Handle.ToInt64(), OriginalLaunchHandleOrigin = "CreateProcessW",
                OriginalProcessHandleOrigin = "Process.GetProcessById_PinnedAfterOwnedCreateProcess",
                Desktop = "Default", InputDesktop = "Default", WindowStation = "WinSta0", Inventory = inventory,
                WordRoots = inventory.Windows.Where(row => row.ClassName == "OpusApp" && row.Visible).ToArray(),
                Vbe = inventory.Windows.Single(row => row.ClassName == "wndclass_desked_gsk" && row.Visible),
                AssemblyPath = path, AssemblyModuleVersionId = (string)status["AssemblyModuleVersionId"], AssemblySha256 = expectedHash,
                ExpectedAssemblyMvid = typeof(VbeSession).Module.ModuleVersionId.ToString("D"),
                NativeRcwReleaseProven = false, HostExitProven = false, Utc = DateTime.UtcNow.ToString("o") };
            using (var stream = new FileStream(Path.Combine(Root, "main-word-placement.json"), FileMode.CreateNew, FileAccess.Write))
            using (var writer = new StreamWriter(stream, new UTF8Encoding(false))) writer.Write(new JavaScriptSerializer().Serialize(record));
            steps.Add(record);
        }

        internal static string RequireMainWordProduct(IDictionary<string, object> status, int processId,
            string expectedMvid, string expectedHash, Guid localMvid, Func<string, string> hash)
        {
            Guid pin; object pid, path, mvid;
            if (status == null || hash == null || processId <= 0 || !Guid.TryParseExact(expectedMvid, "D", out pin) || pin != localMvid ||
                expectedHash == null || expectedHash.Length != 64 || expectedHash.Any(value => !Uri.IsHexDigit(value)) ||
                !status.TryGetValue("HostProcessId", out pid) || !(pid is int) || (int)pid != processId ||
                !status.TryGetValue("AssemblyPath", out path) || !(path is string) ||
                !status.TryGetValue("AssemblyModuleVersionId", out mvid) || !string.Equals(mvid as string, expectedMvid, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The actual Word status must match the frozen candidate and original host.");
            string canonical = ExcelOwnedBootstrapPlan.RequireLocalAbsolutePath((string)path);
            if (!string.Equals(canonical, (string)path, StringComparison.OrdinalIgnoreCase) ||
                !string.Equals(hash(canonical), expectedHash, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The actual loaded Word product path/hash differs from the candidate.");
            return canonical;
        }
    }
}
