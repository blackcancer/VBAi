using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Integration
{
    internal sealed partial class OfficeVbeFixture
    {
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

        internal void QualifyWordGitProjectIsolation()
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
                var firstGit = new VbaGitProject(() => resolve(firstPath), firstPath, readPath);
                var secondGit = new VbaGitProject(() => resolve(secondPath), secondPath, readPath);
                var firstCode = Data("read_module", "Project", firstPath, "Module", "QualificationMarker");
                var secondCode = Data("read_module", "Project", secondPath, "Module", "QualificationMarker");
                GitPhase("First production Git capture", () => {
                    var snapshot = firstGit.Capture();
                    AssertOwnedMarker(snapshot, "FirstOwnedMarker", "SecondOwnedMarker");
                    Assert.IsFalse(snapshot.Manifest.Components.Any(c => c.Type == 3));
                });
                GitPhase("Second production Git capture", () => {
                    var snapshot = secondGit.Capture();
                    AssertOwnedMarker(snapshot, "SecondOwnedMarker", "FirstOwnedMarker");
                    Assert.IsFalse(snapshot.Manifest.Components.Any(c => c.Type == 3));
                });
                // Distinct diagnostic routes and fresh destinations, never a retry of a failed destination.
                GitPhase("Word module export through owning-process bridge", () => {
                    string destination = Path.Combine(Root, "bridge-marker.bas");
                    Assert.IsFalse(File.Exists(destination));
                    var state = Data("component_properties", "Project", firstPath, "Module", "QualificationMarker");
                    Data("export_component", "Project", firstPath, "Module", "QualificationMarker", "Path", destination,
                        "ExpectedComponentVersion", state["Version"]);
                    StringAssert.Contains(File.ReadAllText(destination), "FirstOwnedMarker");
                });
                GitPhase("Word module export through external STA", () => {
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
                // Exercise the real in-process bridge selection data consumed by chat and Immediate guards.
                GitPhase("Word canonical native selection", () => {
                    Data("select_code", "Project", secondPath, "Module", "QualificationMarker", "StartLine", 1, "ExpectedSha256", secondCode["Sha256"]);
                    var selection = Data("debug_state", "Project", secondPath);
                    Assert.AreEqual(secondPath, selection["SelectedHostPath"]);
                    Assert.AreEqual(commonName, selection["SelectedProject"]);
                });
                steps.Add(new { Scope = "Word Git exact-path diagnostic",
                    FirstPath = firstPath, SecondPath = secondPath, CommonProjectName = commonName,
                    GitExecution = "Production Git adapter on the test STA, with explicit owned-PID native Word path probe",
                    ChatEvidence = "Real bridge canonical scope/selection fields; actual chat UI opening is not claimed",
                    RemoteOperations = 0, MacroExecutions = 0, UserForms = 0 });
                GitPhase("Word Git refuses stale SaveAs binding", () => {
                    ((dynamic)secondDocument).SaveAs2(renamedPath, 13); // One owned rename, no user file overwrite.
                    // Old selector must fail before Capture reaches any Export. No successful-path recapture is attempted.
                    Assert.ThrowsException<InvalidOperationException>(() => secondGit.Capture());
                    Assert.AreEqual(false, Response("list_modules", "Project", secondPath)["Ok"]);
                    Assert.IsTrue(probe.SameProject(secondProject, resolve(renamedPath)));
                    Assert.AreEqual(secondCode["Sha256"], Data("read_module", "Project", renamedPath, "Module", "QualificationMarker")["Sha256"]);
                });
                GitPhase("Other same-name Word document source preserved", () => {
                    Assert.AreEqual(firstCode["Sha256"], Data("read_module", "Project", firstPath, "Module", "QualificationMarker")["Sha256"]);
                    Assert.AreEqual(firstPath, readPath(firstProject));
                });
            }
            finally
            {
                if (secondDocument != null)
                    try { ((dynamic)secondDocument).Close(0); }
                    catch (Exception error) { Failures.Add("Owned second Word document close: " + error.Message); }
                Release(editor); Release(secondProject); Release(firstProject); Release(secondDocument); Release(documents);
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
