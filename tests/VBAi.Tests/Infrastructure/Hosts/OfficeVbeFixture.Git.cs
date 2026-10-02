using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Integration
{
    internal sealed partial class OfficeVbeFixture
    {
        private List<object> uncertainWordGitIsolationLeases;
        // The test process is not WINWORD. Supply the owned Word application's PID explicitly;
        // all document enumeration, native HWND checks and IUnknown comparisons remain production code.
        private sealed class OwnedWordPathProbe : VbeProjectComponents.IOtherHostProbe
        {
            private readonly OfficeVbeFixture fixture;
            private readonly VbeProjectComponents.NativeOtherHostProbe native =
                new VbeProjectComponents.NativeOtherHostProbe { ReadHostKind = () => "Word" };
            internal OwnedWordPathProbe(OfficeVbeFixture fixture) { this.fixture = fixture; }
            public string HostKind => "Word";
            public int CurrentProcessId => fixture.ProcessId;
            public object Application() => fixture.application;
            public uint ApplicationProcessId(object application) => native.ApplicationProcessId(application);
            public IList<object> Documents(object application) => native.Documents(application);
            public object DocumentProject(object document) => native.DocumentProject(document);
            public bool SameProject(object a, object b) => native.SameProject(a, b);
            public VbeProjectComponents.OtherHostDocumentState State(object document) => native.State(document);
            public bool FileExists(string path) => File.Exists(path);
            public bool DirectoryExists(string path) => Directory.Exists(path);
            public long FileLength(string path) => new FileInfo(path).Length;
            public void Save(object document, bool saveAs, string destination, int format) =>
                throw new AssertFailedException("The Git path probe is strictly read-only.");
        }

        internal void QualifyWordCanonicalPathSelection()
            => QualifyWordGitProjectIsolation(false);

        /// <summary>Historical external-STA Capture/export diagnostic, enabled only by its dedicated native opt-in.</summary>
        internal void DiagnoseWordGitCaptureExports()
            => QualifyWordGitProjectIsolation(true);

        private void QualifyWordGitProjectIsolation(bool captureDiagnostic)
        {
            Assert.AreEqual("Word", Kind); Assert.IsTrue(owned);
            Assert.IsFalse(ownedProcess.HasExited);
            string firstPath = DocumentPath;
            string secondPath = Path.Combine(Root, "Second.docm");
            string renamedPath = Path.Combine(Root, "SecondRenamed.docm");
            Assert.IsFalse(File.Exists(secondPath)); Assert.IsFalse(File.Exists(renamedPath));
            object documents = null, secondDocument = null, firstProject = null, secondProject = null, editor = null;
            try
            {
                RequireOwnedDocument();
                var probe = new OwnedWordPathProbe(this);
                Assert.AreEqual((uint)ProcessId, probe.ApplicationProcessId(application));
                documents = ((dynamic)application).Documents;
                secondDocument = ((dynamic)documents).Add();
                ((dynamic)secondDocument).SaveAs2(secondPath, 13); // One save of a newly owned document only.
                firstProject = ((dynamic)document).VBProject;
                secondProject = ((dynamic)secondDocument).VBProject;
                string commonName = (string)((dynamic)firstProject).Name;
                Assert.AreEqual(commonName, (string)((dynamic)secondProject).Name, "This regression requires the actual duplicate Word project names.");
                editor = ((dynamic)application).VBE;
                Func<object, string> readPath = project => VbeProjectHostPath.Read(project, probe);
                Func<string, object> resolve = path => (object)VbeProjectResolver.Resolve((dynamic)editor, path, readPath);
                Assert.IsTrue(probe.SameProject(firstProject, resolve(firstPath)));
                Assert.IsTrue(probe.SameProject(secondProject, resolve(secondPath)));
                var projects = Items("list_projects");
                foreach (string path in new[] { firstPath, secondPath })
                {
                    var row = projects.Single(p => string.Equals(Convert.ToString(p["HostPath"]), path, StringComparison.OrdinalIgnoreCase));
                    Assert.AreEqual(commonName, row["Name"]);
                    Assert.AreEqual(path, VbeProjectHostPath.FromFields(row));
                }
                Assert.AreEqual(false, Response("list_modules", "Project", commonName)["Ok"], "Duplicate names must remain ambiguous.");
                PrepareGitMarker(firstPath, "FirstOwnedMarker");
                PrepareGitMarker(secondPath, "SecondOwnedMarker");
                var secondGit = new VbaGitProject(() => resolve(secondPath), secondPath, readPath);
                // A retained exact project also exercises VbaGitProject.CheckedProject after SaveAs,
                // without entering Capture/Export or reaching CodePane mutation on refusal.
                var retainedSecondGit = new VbaGitProject(() => secondProject, secondPath, readPath);
                var firstCode = Data("read_module", "Project", firstPath, "Module", "QualificationMarker");
                var secondCode = Data("read_module", "Project", secondPath, "Module", "QualificationMarker");
                if (captureDiagnostic)
                {
                    var firstGit = new VbaGitProject(() => resolve(firstPath), firstPath, readPath);
                    GitPhase("Diagnostic only: first external-STA production Git Capture", () => {
                        var snapshot = firstGit.Capture();
                        AssertOwnedMarker(snapshot, "FirstOwnedMarker", "SecondOwnedMarker");
                        Assert.IsFalse(snapshot.Manifest.Components.Any(c => c.Type == 3));
                    });
                    GitPhase("Diagnostic only: second external-STA production Git Capture", () => {
                        var snapshot = secondGit.Capture();
                        AssertOwnedMarker(snapshot, "SecondOwnedMarker", "FirstOwnedMarker");
                        Assert.IsFalse(snapshot.Manifest.Components.Any(c => c.Type == 3));
                    });
                    // Separate fresh destinations are diagnostics, never a retry of a failed Capture export.
                    GitPhase("Diagnostic only: Word module export through owning-process bridge", () => {
                        string destination = Path.Combine(Root, "bridge-marker.bas");
                        Assert.IsFalse(File.Exists(destination));
                        var state = Data("component_properties", "Project", firstPath, "Module", "QualificationMarker");
                        Data("export_component", "Project", firstPath, "Module", "QualificationMarker", "Path", destination,
                            "ExpectedComponentVersion", state["Version"]);
                        StringAssert.Contains(File.ReadAllText(destination), "FirstOwnedMarker");
                    });
                    GitPhase("Diagnostic only: Word module export through external STA", () => {
                        string destination = Path.Combine(Root, "external-marker.bas");
                        Assert.IsFalse(File.Exists(destination));
                        object collection = null, component = null;
                        try
                        {
                            collection = ((dynamic)firstProject).VBComponents;
                            component = ((dynamic)collection).Item("QualificationMarker");
                            ((dynamic)component).Export(destination);
                            StringAssert.Contains(File.ReadAllText(destination), "FirstOwnedMarker");
                        }
                        finally { Release(component); Release(collection); }
                    });
                }
                // Exercise the real in-process bridge selection data consumed by chat and Immediate guards.
                PathPhase("Word canonical native selection", () => {
                    Data("select_code", "Project", secondPath, "Module", "QualificationMarker", "StartLine", 1, "ExpectedSha256", secondCode["Sha256"]);
                    var selection = Data("debug_state", "Project", secondPath);
                    Assert.AreEqual(secondPath, selection["SelectedHostPath"]);
                    Assert.AreEqual(commonName, selection["SelectedProject"]);
                });
                steps.Add(new { Scope = "Word canonical path, selection and stale SaveAs qualification",
                    FirstPath = firstPath, SecondPath = secondPath, CommonProjectName = commonName,
                    CaptureExportDiagnosticOptIn = captureDiagnostic,
                    GitExecution = captureDiagnostic ? "Historical external-STA Capture/export diagnostic" : "No Git Capture/export",
                    ChatEvidence = "Real bridge canonical scope/selection fields; actual chat UI opening is not claimed",
                    RemoteOperations = 0, MacroExecutions = 0, UserForms = 0 });
                PathPhase("Word Git refuses stale SaveAs binding before CodePane access", () => {
                    RequireActiveWordGitDocument(secondDocument, secondProject, secondPath, probe);
                    try { ((dynamic)secondDocument).SaveAs2(renamedPath, 13); } // One owned rename, no user file overwrite.
                    catch { NativeExecutionUnsettled = true; throw; } // Unknown SaveAs outcome retains the original host.
                    // Both the old canonical resolver and the retained-project adapter must refuse
                    // before OpenModule can reach CodePane.Show; no Capture/Export is involved.
                    var oldSelector = Assert.ThrowsException<InvalidOperationException>(() => secondGit.OpenModule("QualificationMarker"));
                    StringAssert.Contains(oldSelector.Message, "Project selector is absent or ambiguous:");
                    var changedBinding = Assert.ThrowsException<InvalidOperationException>(() => retainedSecondGit.OpenModule("QualificationMarker"));
                    Assert.AreEqual(UiText.Get("The linked document changed. Reopen GitHub integration."), changedBinding.Message);
                    Assert.AreEqual(false, Response("list_modules", "Project", secondPath)["Ok"]);
                    Assert.IsTrue(probe.SameProject(secondProject, resolve(renamedPath)));
                    Assert.AreEqual(renamedPath, readPath(secondProject), true);
                    Assert.AreEqual(secondCode["Sha256"], Data("read_module", "Project", renamedPath, "Module", "QualificationMarker")["Sha256"]);
                });
                PathPhase("Other same-name Word document source preserved", () => {
                    Assert.AreEqual(firstCode["Sha256"], Data("read_module", "Project", firstPath, "Module", "QualificationMarker")["Sha256"]);
                    Assert.AreEqual(firstPath, readPath(firstProject));
                });
            }
            finally
            {
                if (NativeExecutionUnsettled)
                {
                    uncertainWordGitIsolationLeases = new List<object> { editor, secondProject, firstProject,
                        secondDocument, documents };
                    steps.Add(new { Scenario = "Word Git uncertain native mutation retained", ProcessId,
                        CloseAttempted = false, ReplayAttempts = 0 });
                }
                else
                {
                    if (secondDocument != null)
                        try { ((dynamic)secondDocument).Close(0); }
                        catch (Exception error) { Failures.Add("Owned second Word document close: " + error.Message); }
                    Release(editor); Release(secondProject); Release(firstProject); Release(secondDocument); Release(documents);
                }
            }
        }

        private void RequireActiveWordGitDocument(object expectedDocument, object expectedProject,
            string expectedPath, OwnedWordPathProbe probe)
        {
            RequireUsableOwnedHost();
            Assert.IsTrue(owned && ownedProcess != null && !ownedProcess.HasExited);
            Assert.AreEqual((uint)ProcessId, probe.ApplicationProcessId(application));
            object active = null, window = null, project = null;
            IntPtr expectedUnknown = IntPtr.Zero, activeUnknown = IntPtr.Zero;
            try
            {
                active = ((dynamic)application).ActiveDocument;
                expectedUnknown = Marshal.GetIUnknownForObject(expectedDocument);
                activeUnknown = Marshal.GetIUnknownForObject(active);
                Assert.AreEqual(expectedUnknown, activeUnknown, "The active Word document is not the selected owned SaveAs document.");
                Assert.AreEqual(expectedPath, Path.GetFullPath((string)((dynamic)active).FullName), true);
                window = ((dynamic)application).ActiveWindow;
                IntPtr handle = new IntPtr(Convert.ToInt64(((dynamic)window).hWnd));
                uint pid; GetWindowThreadProcessId(handle, out pid);
                Assert.AreEqual((uint)ProcessId, pid, "The selected Word window belongs to another process.");
                project = ((dynamic)active).VBProject;
                Assert.IsTrue(probe.SameProject(expectedProject, project));
                Assert.AreEqual(expectedPath, VbeProjectHostPath.Read(project, probe), true);
            }
            finally
            {
                if (activeUnknown != IntPtr.Zero) Marshal.Release(activeUnknown);
                if (expectedUnknown != IntPtr.Zero) Marshal.Release(expectedUnknown);
                // ActiveDocument/VBProject getters can return the retained second document's RCWs.
                // Balance only these acquisitions; FinalReleaseComObject would invalidate the retained leases.
                if (project != null && Marshal.IsComObject(project)) Marshal.ReleaseComObject(project);
                Release(window);
                if (active != null && Marshal.IsComObject(active)) Marshal.ReleaseComObject(active);
            }
        }

        private void PathPhase(string name, Action action)
        {
            try { action(); steps.Add(new { Scenario = name, Result = "PASS" }); }
            catch (Exception error)
            {
                Failures.Add(name + ": " + error);
                steps.Add(new { Scenario = name, Result = "FAIL", ExceptionType = error.GetType().FullName,
                    HResult = "0x" + unchecked((uint)error.HResult).ToString("X8"), Error = error.ToString() });
                throw;
            }
        }

        private void GitPhase(string name, Action action)
        {
            try { action(); steps.Add(new { Scenario = name, Result = "PASS" }); }
            catch (Exception error)
            {
                Failures.Add(name + ": " + error.GetType().Name + " / 0x" + unchecked((uint)error.HResult).ToString("X8") + ": " + error.Message);
                steps.Add(new { Scenario = name, Result = "FAIL", ExceptionType = error.GetType().FullName,
                    HResult = "0x" + unchecked((uint)error.HResult).ToString("X8"), Error = error.ToString() });
            }
        }

        private void PrepareGitMarker(string path, string marker)
        {
            Data("create_module", "Project", path, "Module", "QualificationMarker", "ExpectedMode", 2);
            var before = Data("read_module", "Project", path, "Module", "QualificationMarker");
            int lines = Convert.ToInt32(Data("component_properties", "Project", path, "Module", "QualificationMarker")["CodeLines"]);
            Data("replace_lines", "Project", path, "Module", "QualificationMarker", "StartLine", 1, "Count", lines,
                "Text", "' " + marker + "\r\n' Synthetic qualification source; never executed.\r\n" + Convert.ToString(before["Code"]), "ExpectedSha256", before["Sha256"]);
        }
        private static void AssertOwnedMarker(VbaGitSnapshot snapshot, string expected, string forbidden)
        {
            string text = VbaGitSnapshot.Utf8.GetString(snapshot.Files["QualificationMarker.bas"]);
            StringAssert.Contains(text, expected); Assert.IsFalse(text.Contains(forbidden));
        }
    }
}
