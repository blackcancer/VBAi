using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Integration
{
    internal sealed partial class OfficeVbeFixture
    {
        // Observations only: neither seam adopts a discovered process or changes the COM application.
        internal Func<IntPtr, uint> ReadPublisherWindowOwner = window => {
            uint pid; GetWindowThreadProcessId(window, out pid); return pid;
        };
        internal Func<object[]> ReadPublisherProcessCanary = ObservePublisherProcesses;
        internal Func<bool> ReadPublisherOwnedProcessAlive;

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct PublisherProcessEntry
        {
            public uint Size, Usage, ProcessId;
            public IntPtr DefaultHeapId;
            public uint ModuleId, ThreadCount, ParentProcessId;
            public int PriorityClass;
            public uint Flags;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)] public string ImageName;
        }
        [DllImport("kernel32.dll", SetLastError = true)] private static extern IntPtr CreateToolhelp32Snapshot(uint flags, uint processId);
        [DllImport("kernel32.dll", EntryPoint = "Process32FirstW", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern bool ReadPublisherProcessFirst(IntPtr snapshot, ref PublisherProcessEntry entry);
        [DllImport("kernel32.dll", EntryPoint = "Process32NextW", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern bool ReadPublisherProcessNext(IntPtr snapshot, ref PublisherProcessEntry entry);
        [DllImport("kernel32.dll", SetLastError = true)] private static extern bool CloseHandle(IntPtr handle);

        private void RecordPublisherActivationCanary()
        {
            steps.Add(new { PublisherStartupCanary = "AfterActivationBeforeNewDocument", ProcessId,
                ObservedProcesses = ReadPublisherProcessCanary(), AutomaticRebind = false });
            FlushAdapterEvidence();
        }

        private sealed class PublicationWindow
        {
            public long Handle { get; set; }
            public uint ProcessId { get; set; }
            public string Caption { get; set; }
        }

        private sealed class PublicationIdentity
        {
            public string FullName { get; set; }
            public bool Saved { get; set; }
            public string IUnknown { get; set; }
            public PublicationWindow Window { get; set; }
        }

        /// <summary>Requires one publication belonging to the retained process; wrapper identity alone is not a document association.</summary>
        internal void RequirePublisherPublication(string phase, bool requireSavedPath)
        {
            commandContainment.RequireTerminal();
            if (publisherBootstrap != null) RecheckPrivatePublisherBootstrap();
            object active = null, collection = null, sole = null, appWindow = null;
            var proof = new Dictionary<string, object> { ["PublisherStartupCanary"] = phase, ["ProcessId"] = ProcessId,
                ["ExpectedDocumentPath"] = DocumentPath, ["RequireSavedPath"] = requireSavedPath,
                ["CaptureState"] = "Reading", ["AutomaticRebind"] = false };
            steps.Add(proof); FlushAdapterEvidence();
            try
            {
                proof["ObservedProcesses"] = ReadPublisherProcessCanary(); FlushAdapterEvidence();
                bool alive = ReadPublisherOwnedProcessAlive != null ? ReadPublisherOwnedProcessAlive() :
                    owned && ownedProcess != null && !ownedProcess.HasExited;
                proof["RetainedProcessAlive"] = alive;
                Assert.IsTrue(alive, "The retained Publisher process handle must still be alive; PID reuse cannot establish ownership.");
                var retained = CapturePublication(document);
                proof["RetainedDocument"] = retained;
                active = ((dynamic)application).ActiveDocument;
                var selected = CapturePublication(active); proof["ActiveDocument"] = selected;
                appWindow = ((dynamic)application).ActiveWindow;
                var applicationWindow = CapturePublicationWindow(appWindow); proof["ApplicationWindow"] = applicationWindow;
                collection = ((dynamic)application).Documents;
                int count = Convert.ToInt32(((dynamic)collection).Count); proof["DocumentCount"] = count;
                Assert.AreEqual(1, count, "Publisher must expose exactly one publication; recovered documents are retained without mutation.");
                sole = ((dynamic)collection)[1];
                var only = CapturePublication(sole); proof["SoleDocument"] = only;
                RequirePublisherDocumentIdentities(retained.IUnknown, selected.IUnknown, only.IUnknown);
                var windows = new[] { retained.Window, selected.Window, only.Window, applicationWindow };
                Assert.IsTrue(windows.All(w => w.Handle != 0 && w.ProcessId == (uint)ProcessId),
                    "The document and application windows must belong to the retained Publisher PID. A child is observed, never adopted.");
                Assert.IsTrue(windows.All(w => w.Handle == retained.Window.Handle), "Publisher document/application windows disagree.");
                foreach (var identity in new[] { retained, selected, only })
                {
                    if (requireSavedPath)
                    {
                        Assert.IsTrue(SamePublicationPath(identity.FullName, DocumentPath), "The publication is not the exact disposable file.");
                        Assert.IsTrue(identity.Saved, "The disposable publication bootstrap must already be saved.");
                    }
                    else
                        Assert.AreEqual(retained.FullName, identity.FullName, true, "The returned new publication is not the active sole publication.");
                }
                if (publisherBootstrap != null)
                    publisherBootstrap.ConfirmNativePublication(() => {
                        var nativeWindows = new List<PublisherOwnerWindow>();
                        foreach (var window in windows)
                        {
                            var native = ReadPrivatePublisherOwnerWindow(new IntPtr(window.Handle));
                            if (native.Process != (uint)ProcessId || native.RootProcess != (uint)ProcessId || native.Thread == 0 || native.RootThread == 0 || native.Root == IntPtr.Zero)
                                throw new InvalidOperationException("Publisher publication has no exact original-process native root/thread proof.");
                            nativeWindows.Add(native);
                        }
                        proof["NativeWindows"] = PublisherOwnershipProbe.EvidenceValue(nativeWindows.ToArray());
                        RecheckPrivatePublisherBootstrap();
                        proof["CanonicalApplicationIUnknown"] = ReadPrivatePublisherBootstrapIdentity();
                    });
                proof["Verified"] = true;
                proof["CanonicalDocumentIdentityVerified"] = true;
            }
            catch (Exception error)
            {
                proof["Verified"] = false; proof["Error"] = error.ToString();
                RetainUncertainOffice();
                throw;
            }
            finally
            {
                // Each getter acquisition is balanced once, even when several getters share one retained RCW.
                BalancePublicationGetter(sole); BalancePublicationGetter(collection);
                BalancePublicationGetter(appWindow); BalancePublicationGetter(active);
                proof["CaptureState"] = "Terminal"; FlushAdapterEvidence();
            }
        }

        private PublicationIdentity CapturePublication(object value)
        {
            Assert.IsNotNull(value, "Publisher did not return a document.");
            object window = null;
            try
            {
                window = ((dynamic)value).ActiveWindow;
                return new PublicationIdentity { FullName = (string)((dynamic)value).FullName, Saved = (bool)((dynamic)value).Saved,
                    IUnknown = PublicationComIdentity(value), Window = CapturePublicationWindow(window) };
            }
            finally { BalancePublicationGetter(window); }
        }

        private PublicationWindow CapturePublicationWindow(object window)
        {
            Assert.IsNotNull(window, "Publisher did not return a document window.");
            long handle = Convert.ToInt64(((dynamic)window).Hwnd);
            return new PublicationWindow { Handle = handle, ProcessId = ReadPublisherWindowOwner(new IntPtr(handle)),
                Caption = (string)((dynamic)window).Caption };
        }

        private static string PublicationComIdentity(object value)
        {
            if (!Marshal.IsComObject(value)) return "managed-fake";
            IntPtr unknown = Marshal.GetIUnknownForObject(value);
            try { return unknown.ToInt64().ToString("X"); }
            finally { Marshal.Release(unknown); }
        }

        private static void BalancePublicationGetter(object value)
        { if (value != null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value); }

        private static bool SamePublicationPath(string first, string second)
        {
            if (string.IsNullOrWhiteSpace(first) || string.IsNullOrWhiteSpace(second) || !Path.IsPathRooted(first) || !Path.IsPathRooted(second)) return false;
            try { return string.Equals(Path.GetFullPath(first), Path.GetFullPath(second), StringComparison.OrdinalIgnoreCase); }
            catch (ArgumentException) { return false; }
            catch (NotSupportedException) { return false; }
            catch (PathTooLongException) { return false; }
        }

        private static object Field(IDictionary<string, object> values, string key)
        { object value; return values.TryGetValue(key, out value) ? value : null; }

        /// <summary>Binds only an exactly mapped, selected, available Publisher project after read-only identity verification.</summary>
        private void BindPublisherStartupProject(IDictionary<string, object>[] projects)
        {
            Project = null;
            try
            {
                RequirePublisherPublication("BeforeProjectBinding", true);
                var matched = projects.Where(p => SamePublicationPath(Field(p, "FileName") as string, DocumentPath) ||
                    SamePublicationPath(Field(p, "HostPath") as string, DocumentPath)).ToArray();
                // Publisher can expose no VBIDE/catalog path at all. Only a sole pathless project
                // is provisional: the adapter must independently prove the exact native document below.
                bool solePathless = matched.Length == 0 && projects.Length == 1 &&
                    string.IsNullOrWhiteSpace(Field(projects[0], "FileName") as string) &&
                    string.IsNullOrWhiteSpace(Field(projects[0], "HostPath") as string);
                if (solePathless) matched = projects;
                Assert.AreEqual(1, matched.Length, "No unique exact Publisher project mapping; generic Project fallback is forbidden.");
                var candidate = matched[0];
                string inventoryPath = Field(candidate, "FileName") as string;
                Assert.IsTrue(string.IsNullOrWhiteSpace(inventoryPath) || SamePublicationPath(inventoryPath, DocumentPath),
                    "Publisher inventory selected a recovered or foreign VBIDE path.");
                string inventoryHostPath = Field(candidate, "HostPath") as string;
                Assert.IsTrue(string.IsNullOrWhiteSpace(inventoryHostPath) || SamePublicationPath(inventoryHostPath, DocumentPath),
                    "Publisher inventory mapped a foreign host path.");
                string name = Field(candidate, "Name") as string;
                Assert.IsFalse(string.IsNullOrWhiteSpace(name), "The mapped Publisher project has no name.");
                // Use an explicit provisional selector without binding the baseline before verification succeeds.
                string selector = SamePublicationPath(Field(candidate, "FileName") as string, DocumentPath) ? DocumentPath : name;
                var persistence = Data("project_persistence_status", "Project", selector);
                var selection = Data("debug_state", "Project", selector);
                steps.Add(new { PublisherStartupProject = candidate, Persistence = persistence, Selection = selection,
                    ProvisionalSelector = selector, SolePathlessProject = solePathless });
                RequirePublisherPersistenceProof(persistence, selector);
                Assert.AreEqual(2, Convert.ToInt32(Field(selection, "Mode")), "The selected publication must be in design mode.");
                if (Field(selection, "SelectedProject") == null)
                {
                    // State.SelectedProject describes an active CodePane, not ActiveVBProject.
                    // A fresh publication may have no pane. Its exact singleton native COM/VBE
                    // association is already verified by OtherHostPersistence/SolePublisherProject.
                    Assert.AreEqual(1, projects.Length, "An absent code pane cannot disambiguate more than one native project.");
                    Assert.AreEqual(name, Field(selection, "Project") as string, true, "The no-pane response belongs to another requested project.");
                    foreach (string key in new[] { "SelectedProject", "SelectedProjectPath", "SelectedHostPath", "ActiveModule", "Selection" })
                        Assert.IsTrue(selection.ContainsKey(key) && selection[key] == null, "The native pane context is partial, foreign or failed: " + key);
                    var currentProjects = Items("list_projects");
                    Assert.AreEqual(1, currentProjects.Length, "The native project catalog changed before binding the no-pane publication.");
                    var currentProject = currentProjects[0];
                    Assert.AreEqual(name, Field(currentProject, "Name") as string, true);
                    Assert.AreEqual(2, Convert.ToInt32(Field(currentProject, "Mode")));
                    foreach (string key in new[] { "FileName", "HostPath" })
                    {
                        string currentPath = Field(currentProject, key) as string;
                        Assert.IsTrue(string.IsNullOrWhiteSpace(currentPath) || SamePublicationPath(currentPath, DocumentPath),
                            "The no-pane native project acquired a foreign path: " + key);
                    }
                    var finalPersistence = Data("project_persistence_status", "Project", selector);
                    steps.Add(new { PublisherNoCodePaneAssociation = true, CurrentProjects = currentProjects, FinalPersistence = finalPersistence,
                        Proof = "Exact singleton native document/VBE association; no code-pane selection is inferred" });
                    RequirePublisherPersistenceProof(finalPersistence, selector);
                }
                else Assert.AreEqual(name, Field(selection, "SelectedProject") as string, true, "The active VBE project is not the mapped publication.");
                string selectedHostPath = Field(selection, "SelectedHostPath") as string;
                Assert.IsTrue(((solePathless || Field(selection, "SelectedProject") == null) && string.IsNullOrWhiteSpace(selectedHostPath)) || SamePublicationPath(selectedHostPath, DocumentPath),
                    "VBE selected a recovered or foreign publication.");
                string selectedPath = Field(selection, "SelectedProjectPath") as string;
                Assert.IsTrue(string.IsNullOrWhiteSpace(selectedPath) || SamePublicationPath(selectedPath, DocumentPath), "Selected VBIDE path disagrees with the publication.");
                RequirePublisherPublication("AfterProjectBindingReadback", true);
                Project = selector;
            }
            catch (Exception error)
            {
                RetainUncertainOffice();
                steps.Add(new { PublisherStartupBindingError = error.ToString(), BaselineBound = false });
                FlushAdapterEvidence(); throw;
            }
        }

        private void RequirePublisherPersistenceProof(IDictionary<string, object> proof, string selector)
        {
            Assert.AreEqual(selector, Field(proof, "Project") as string, true, "The adapter proof did not target the exact provisional project selector.");
            Assert.AreEqual(true, Field(proof, "HostAvailable"), "Publisher's project host is unavailable.");
            Assert.AreEqual(true, Field(proof, "IdentityVerified"), "Publisher's project identity was not verified.");
            Assert.AreEqual(ProcessId, Convert.ToInt32(Field(proof, "OwnerProcessId")), "Publisher project belongs to another PID.");
            Assert.AreEqual("Publisher", Field(proof, "Host") as string, true);
            Assert.IsTrue(SamePublicationPath(Field(proof, "HostPath") as string, DocumentPath), "The adapter did not map the exact disposable publication.");
        }

        private static object[] ObservePublisherProcesses()
        {
            var parents = new Dictionary<uint, uint>();
            IntPtr snapshot = CreateToolhelp32Snapshot(2, 0);
            int snapshotError = 0;
            if (snapshot == new IntPtr(-1)) snapshotError = Marshal.GetLastWin32Error();
            else
            {
                try
                {
                    var entry = new PublisherProcessEntry { Size = (uint)Marshal.SizeOf(typeof(PublisherProcessEntry)) };
                    if (ReadPublisherProcessFirst(snapshot, ref entry))
                        do { parents[entry.ProcessId] = entry.ParentProcessId; } while (ReadPublisherProcessNext(snapshot, ref entry));
                    else snapshotError = Marshal.GetLastWin32Error();
                }
                finally { CloseHandle(snapshot); }
            }
            var processes = Process.GetProcessesByName("MSPUB");
            try
            {
                return processes.Select(process => {
                    var record = new Dictionary<string, object> { ["ProcessId"] = process.Id, ["ObservedOnly"] = true, ["Adopted"] = false };
                    uint parent; record["ParentProcessId"] = parents.TryGetValue((uint)process.Id, out parent) ? (object)parent : null;
                    record["ParentObservationError"] = snapshotError == 0 ? null : (object)snapshotError;
                    try { record["StartTimeUtc"] = process.StartTime.ToUniversalTime().ToString("O"); record["ImagePath"] = process.MainModule.FileName; }
                    catch (Exception error) { record["ObservationError"] = error.Message; }
                    return (object)record;
                }).ToArray();
            }
            finally { foreach (var process in processes) process.Dispose(); }
        }
    }
}
