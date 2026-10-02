using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Web.Script.Serialization;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Integration
{
    internal sealed partial class ExcelVbeFixture
    {
        // Keep native identity alive across readbacks even when temporary shared RCWs are released.
        private IntPtr embeddedGitProjectIdentity;
        internal sealed class EmbeddedGitScope
        {
            internal string Path, Cache, State, Marker, References;
            internal IntPtr VbeHandle;
            internal uint ThreadId;
            internal Dictionary<string, string> Code;
            internal Dictionary<string, int> Types;
            internal VbaGitSnapshot Baseline;
            internal string Layout;
            internal IDictionary<string, object> NativeLayout, NativeFonts;
            internal EmbeddedGitRepositoryBinding SelectedRepository;
        }

        /// <summary>Prepares only the explicitly owned synthetic workbook; no macro is executed.</summary>
        internal EmbeddedGitScope PrepareEmbeddedGitScope(string marker, Action<bool> pending, Action<object> evidence,
            string layout = null, RootFontObservationManifest.Configuration fontObservation = null)
        {
            string rootFontSeedProfile = fontObservation?.SeedProfile;
            if (rootFontSeedProfile != null && (layout == null ||
                rootFontSeedProfile != RootFontObservationManifest.SyntheticExplicitArial9 &&
                rootFontSeedProfile != RootFontObservationManifest.RetainedSyntheticTahoma825))
                throw new InvalidOperationException("A synthetic root font seed requires one declared layout and profile.");
            bool retained = rootFontSeedProfile == RootFontObservationManifest.RetainedSyntheticTahoma825;
            if (retained && (layout != "LabelButton" || fontObservation?.SourceWorkbook == null))
                throw new InvalidOperationException("The retained source requires its pinned LabelButton configuration.");
            string path = File("EmbeddedGit.xlsm");
            string cache = MacroGitRepository.ScopeDirectory(Path.GetFullPath(path));
            Assert.IsFalse(Directory.Exists(cache), "A fresh workbook must not inherit a previous document's Git cache.");
            Assert.IsFalse(Directory.Exists(MacroGitRepository.ScopeDirectory(Path.GetFullPath(path).ToUpperInvariant())),
                "A fresh workbook must not resolve a legacy document binding.");
            pending(true);
            ((dynamic)application).EnableEvents = false;
            ((dynamic)application).AutomationSecurity = 3;
            object project = null, components = null;
            if (retained)
            {
                evidence(new { Phase = "RetainedCopyIntent", SourceWorkbook = fontObservation.SourceWorkbook,
                    SourceSha256 = RetainedRootFontWorkbook.SourceSha256, Destination = path,
                    BaselineFontSeedProfile = rootFontSeedProfile, CorrelationNonce = marker,
                    SourceMarker = RetainedRootFontWorkbook.SourceMarker });
                OpenRetainedSyntheticWorkbook(fontObservation.SourceWorkbook, path, pending);
                evidence(new { Phase = "RetainedCopyVerified", Destination = path,
                    WorkbookSha256 = RetainedRootFontWorkbook.SourceSha256,
                    SourceSha256 = RetainedRootFontWorkbook.SourceSha256,
                    SourceMarker = RetainedRootFontWorkbook.SourceMarker });
            }
            else
            {
                try
                {
                    project = ((dynamic)workbook).VBProject; components = ((dynamic)project).VBComponents;
                    foreach (var item in new[] { Tuple.Create("EmbeddedModule", 1), Tuple.Create("EmbeddedClass", 2) })
                    {
                        object component = null, code = null;
                        try
                        {
                            component = ((dynamic)components).Add(item.Item2); ((dynamic)component).Name = item.Item1;
                            code = ((dynamic)component).CodeModule;
                            ((dynamic)code).InsertLines(1, "Option Explicit\r\n' " + marker + " " + item.Item1 + "\r\n");
                        }
                        finally { Release(code); Release(component); }
                    }
                }
                finally { Release(components); Release(project); }
                if (layout == null) PrepareGitForm("EmbeddedForm", "Synthetic embedded Git form", marker, path);
                else
                {
                    if (rootFontSeedProfile != null)
                        evidence(new { Phase = "BaselineFontSeedIntent", BaselineFontSeedProfile = rootFontSeedProfile,
                            TargetDescriptorHex = BitConverter.ToString(RootFontObservationManifest.SyntheticArial9Descriptor()).Replace("-", ""),
                            PlannedNativeDeliveryCount = 1 });
                    PrepareGitLayout("EmbeddedForm", layout, path, persistedBaseline: true,
                        rootFontSeedProfile: rootFontSeedProfile);
                }
            }
            pending(false);
            object editor = null, main = null, module = null, moduleCode = null, pane = null;
            try
            {
                if (retained)
                {
                    ObserveEmbeddedVbeRead("Application.VBE", () => editor = ((dynamic)application).VBE, evidence, ProcessId);
                    ObserveEmbeddedVbeRead("VBE.MainWindow", () => main = ((dynamic)editor).MainWindow, evidence, ProcessId);
                }
                else
                {
                    editor = ((dynamic)application).VBE;
                    main = ((dynamic)editor).MainWindow;
                }
                project = ((dynamic)workbook).VBProject; components = ((dynamic)project).VBComponents;
                module = ((dynamic)components).Item("EmbeddedModule"); moduleCode = ((dynamic)module).CodeModule; pane = ((dynamic)moduleCode).CodePane;
                Assert.AreEqual(IntPtr.Zero, embeddedGitProjectIdentity, "Only one embedded Git scope may own this fixture's identity lease.");
                embeddedGitProjectIdentity = Marshal.GetIUnknownForObject(project);
                pending(true);
                ((dynamic)pane).Show(); ((dynamic)editor).ActiveCodePane = pane;
                ((dynamic)pane).SetSelection(1, 1, 1, 1);
                pending(false);
                var hwnd = new IntPtr(Convert.ToInt64(((dynamic)main).HWnd));
                uint pid; uint tid = GetWindowThreadProcessId(hwnd, out pid);
                Assert.AreEqual((uint)ProcessId, pid); Assert.AreNotEqual(0u, tid);
                var scope = new EmbeddedGitScope { Path = path, Marker = retained ? RetainedRootFontWorkbook.SourceMarker : marker,
                    VbeHandle = hwnd, ThreadId = tid,
                    Cache = cache, Layout = layout };
                VbaGitSnapshot preGetter = null;
                if (retained)
                {
                    RequireRetainedActiveProject(path, project, editor);
                    var status = Command("status");
                    Assert.IsTrue(Convert.ToBoolean(status["Ok"]), "The copied workbook lost the owner bridge connection.");
                    RequireMonacoCandidate(typeof(VbeSession).Module.ModuleVersionId,
                        typeof(VbeSession).Module.ModuleVersionId, ProcessId, VbeBridgeClient.Object(status["Data"]));
                    preGetter = RetainedRootFontWorkbook.AttestFirstCapture(
                        () => { VbaGitSnapshot snapshot = null; WithGitProject(path, git => snapshot = git.Capture()); return snapshot; },
                        RetainedRootFontWorkbook.RequirePinnedBaseline,
                        (snapshot, first) => evidence(new { Phase = "RetainedPreGetterCaptureExact", SourceMarker = scope.Marker,
                            CorrelationNonce = marker, RootDescriptorHex = BitConverter.ToString(first).Replace("-", ""),
                            ResourceBytes = snapshot.Files["EmbeddedForm.frx"].Length,
                            SnapshotFiles = EmbeddedGitSnapshotOracle.Describe(snapshot) }));
                }
                if (layout != null)
                {
                    scope.NativeLayout = ReadGitLayout("EmbeddedForm", layout);
                    scope.NativeFonts = ReadGitLayoutFonts("EmbeddedForm", layout);
                    evidence(new { Phase = "PersistedLayoutReadback", Layout = layout, Properties = scope.NativeLayout, Fonts = scope.NativeFonts });
                }
                scope.State = ReadEmbeddedState(scope, out scope.Code, out scope.Types, out scope.References);
                scope.Baseline = ExportEmbeddedBaseline(scope, pending, evidence);
                if (retained)
                {
                    byte[] reopened = RetainedRootFontWorkbook.RequirePinnedBaseline(scope.Baseline);
                    Assert.IsTrue(preGetter.SameAs(scope.Baseline), "The retained baseline changed after scalar readback or bridge export.");
                    Assert.AreEqual(RetainedRootFontWorkbook.SourceSha256, EmbeddedRawHash(path),
                        "The copied workbook changed before menu execution.");
                    evidence(new { Phase = "RetainedBridgeBaselineExact", SourceMarker = scope.Marker,
                        CorrelationNonce = marker, RootDescriptorHex = BitConverter.ToString(reopened).Replace("-", ""),
                        ResourceBytes = scope.Baseline.Files["EmbeddedForm.frx"].Length });
                }
                if (rootFontSeedProfile != null && !retained)
                {
                    var rootForm = scope.Baseline.Manifest.Components.Single(item => item.Name == "EmbeddedForm" && item.Type == 3);
                    byte[] persisted = RootFontObservationManifest.RequireRoot(scope.Baseline.FormFonts(rootForm),
                        "AfterInitialCapture", rootFontSeedProfile);
                    evidence(new { Phase = "BaselineFontSeedReopenedExact", BaselineFontSeedProfile = rootFontSeedProfile,
                        DescriptorHex = BitConverter.ToString(persisted).Replace("-", ""),
                        ResourceBytes = scope.Baseline.Files["EmbeddedForm.frx"].Length });
                }
                VerifyEmbeddedGitState(scope);
                return scope;
            }
            finally { Release(pane); Release(moduleCode); Release(module); Release(components); Release(project); Release(main); Release(editor); }
        }

        /// <summary>Replaces only the fixture's empty owned workbook with one verified disposable copy.</summary>
        private void OpenRetainedSyntheticWorkbook(string source, string path, Action<bool> pending)
        {
            RetainedRootFontWorkbook.RequirePinnedSource(source);
            bool closeIssued = false, blankClosed = false, openIssued = false, openReturned = false;
            try
            {
                closeIssued = true;
                ((dynamic)workbook).Close(false);
                blankClosed = true;
                pending(false);
                object blank = workbook; workbook = null;
                Release(blank);
                RetainedRootFontWorkbook.CopyCreateNew(source, path, Root,
                    RetainedRootFontWorkbook.SourceSha256);
                openIssued = true;
                pending(true);
                workbook = ((dynamic)workbooks).Open(path, 0, false);
                openReturned = true;
                pending(false);
                Assert.AreEqual(path, Convert.ToString(((dynamic)workbook).FullName), true);
                Assert.AreEqual(1, Convert.ToInt32(((dynamic)workbooks).Count));
                Assert.AreEqual(RetainedRootFontWorkbook.SourceSha256, EmbeddedRawHash(path));
            }
            finally
            {
                // A COM Close/Open call may have applied despite an exception.
                if (closeIssued && !blankClosed || openIssued && !openReturned)
                    PreserveForDiagnosticRecovery = true;
            }
        }

        /// <summary>Attests the copied project selected by the owned VBE before its first resource capture.</summary>
        private static void RequireRetainedActiveProject(string path, object project, object editor)
        {
            object active = null;
            Exception primary = null;
            try
            {
                active = ((dynamic)editor).ActiveVBProject;
                if (!VbeProjectHostPath.SameProject(project, active) ||
                    !string.Equals(VbeProjectHostPath.Read(project), path, StringComparison.OrdinalIgnoreCase) ||
                    Convert.ToInt32(((dynamic)project).Mode) != 2 ||
                    Convert.ToInt32(((dynamic)project).Protection) != 0)
                    throw new InvalidOperationException("The retained workbook is not the selected, editable copied project.");
            }
            catch (Exception error) { primary = error; throw; }
            finally
            {
                FormFontRestoration.ReleaseOwnedReferences(new[] { active }, value => {
                    if (value != null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value);
                }, primary);
            }
        }

        private VbaGitSnapshot ExportEmbeddedBaseline(EmbeddedGitScope scope, Action<bool> pending, Action<object> evidence,
            string exportDirectory = "owner-bridge-baseline")
        {
            string directory = File(exportDirectory); Directory.CreateDirectory(directory);
            var files = new Dictionary<string, byte[]>(StringComparer.Ordinal); var manifest = new List<VbaGitComponent>();
            foreach (var item in scope.Types.OrderBy(x => x.Key, StringComparer.Ordinal))
            {
                var component = new VbaGitComponent { Name = item.Key, Type = item.Value };
                if (item.Value == 100) files.Add(component.FileName, VbaGitSnapshot.Utf8.GetBytes(scope.Code[item.Key]));
                else
                {
                    var state = EmbeddedData(new Request { Command = "component_properties", Project = scope.Path, Module = item.Key }, pending, evidence);
                    string path = Path.Combine(directory, component.FileName);
                    EmbeddedData(new Request { Command = "export_component", Project = scope.Path, Module = item.Key,
                        Path = path, ExpectedComponentVersion = Convert.ToString(state["Version"]) }, pending, evidence);
                    var codec = EmbeddedData(new Request { Command = "inspect_code_file", Path = path }, pending, evidence);
                    int codePage = Convert.ToInt32(codec["SystemAnsiCodePage"]);
                    string text = Encoding.GetEncoding(codePage, EncoderFallback.ExceptionFallback, DecoderFallback.ExceptionFallback).GetString(System.IO.File.ReadAllBytes(path));
                    files.Add(component.FileName, VbaGitSnapshot.Utf8.GetBytes(text));
                    evidence(new { Phase = "BaselineCodecAttested", component.Name, CodePage = codePage, Path = path,
                        NativeExportSha256 = EmbeddedRawHash(path), Codec = codec });
                    if (item.Value == 3)
                    {
                        component.HasResources = System.IO.File.Exists(Path.ChangeExtension(path, ".frx"));
                        Assert.IsTrue(component.HasResources, "Independent owner-bridge export must retain the actual form resources.");
                        files.Add(component.Name + ".frx", System.IO.File.ReadAllBytes(Path.ChangeExtension(path, ".frx")));
                        evidence(new { Phase = "BaselineRawResourceRetained", component.Name,
                            NativeResourceSha256 = EmbeddedRawHash(Path.ChangeExtension(path, ".frx")) });
                    }
                }
                manifest.Add(component);
            }
            var result = new VbaGitSnapshot(new VbaGitManifest { Components = manifest.ToArray(), References = scope.References }, files);
            evidence(new { Phase = "IndependentBridgeBaselineVerified", Files = EmbeddedGitSnapshotOracle.Describe(result) });
            return result;
        }

        /// <summary>Records one retained-copy VBE getter on its owner thread without retry or lease transfer.</summary>
        internal static void ObserveEmbeddedVbeRead(string getter, Action read, Action<object> evidence, int processId)
        {
            if (read == null) throw new ArgumentNullException(nameof(read));
            if (evidence == null) throw new ArgumentNullException(nameof(evidence));
            int thread = System.Threading.Thread.CurrentThread.ManagedThreadId;
            evidence(new { Phase = "VbeAccessIntent", Getter = getter, ProcessId = processId, OwnerThread = thread });
            try { read(); }
            catch (Exception error)
            {
                try
                {
                    evidence(new { Phase = "VbeAccessFailed", Getter = getter, ProcessId = processId,
                        OwnerThread = thread, HResult = "0x" + unchecked((uint)error.HResult).ToString("X8"),
                        ErrorType = error.GetType().FullName });
                }
                catch (Exception reportError) { error.Data["VbeAccessEvidenceError"] = reportError.ToString(); }
                throw;
            }
            // The caller has already acquired and retained any returned COM lease.
            // A receipt failure propagates through its existing finally/release path.
            evidence(new { Phase = "VbeAccessReturned", Getter = getter, ProcessId = processId, OwnerThread = thread });
        }

        internal static string EmbeddedRawHash(string path)
        {
            using (var bytes = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var hash = SHA256.Create())
                return BitConverter.ToString(hash.ComputeHash(bytes)).Replace("-", "");
        }

        private IDictionary<string, object> EmbeddedData(Request request, Action<bool> pending, Action<object> evidence)
        {
            evidence(new { Phase = "BaselineRequestIntent", Request = request }); pending(true);
            var reply = Command(request);
            if (reply == null) throw new InvalidOperationException("Baseline bridge delivery is uncertain; no replay or automatic cleanup.");
            pending(false);
            Assert.AreEqual(true, reply["Ok"], "Independent baseline bridge failure: " + Convert.ToString(reply["Error"]));
            return VbeBridgeClient.Object(reply["Data"]);
        }

        /// <summary>Validates the original retained process using only native process/window identity while the menu is modal.</summary>
        internal void RequireEmbeddedProcess(EmbeddedGitScope scope)
        {
            if (!owned || ownedProcess == null || ownedProcess.HasExited || ownedImagePath == null)
                throw new InvalidOperationException("The original owned Excel process is unavailable; no UI action is allowed.");
            ExcelOwnedBootstrapPlan.VerifyAttachedIdentity(ProcessId, Convert.ToString(startupEvidence["HostExecutable"]),
                Convert.ToString(startupEvidence["HostStartedUtc"]), ProcessId, ownedImagePath(), ownedProcess.StartTime.ToUniversalTime().ToString("o"));
            uint pid; uint tid = GetWindowThreadProcessId(scope.VbeHandle, out pid);
            EmbeddedGitUiProtocol.RequireOwner(ProcessId, scope.ThreadId, scope.VbeHandle.ToInt64(), (int)pid, tid, scope.VbeHandle.ToInt64());
        }

        /// <summary>Executes precisely the existing tagged menu button once, on this fixture's STA; returns only after its modal window closes.</summary>
        internal void ExecuteEmbeddedGitMenu(Action<object> evidence)
        {
            object editor = null, bars = null; var matches = new List<object>();
            try
            {
                editor = ((dynamic)application).VBE; bars = ((dynamic)editor).CommandBars;
                foreach (object bar in (dynamic)bars)
                {
                    try { if (Convert.ToInt32(((dynamic)bar).Type) == 1) FindEmbeddedGitButton(bar, matches, 0); }
                    finally { Release(bar); }
                }
                Assert.AreEqual(1, matches.Count, "Only one exact VBAi.GitHub button is allowed.");
                dynamic button = matches[0]; Assert.AreEqual("VBAi.GitHub", (string)button.Tag); Assert.IsTrue((bool)button.Enabled);
                evidence(new { Phase = "MenuExecuteIntent", Tag = (string)button.Tag, ControlId = (int)button.Id,
                    Caption = (string)button.Caption, OwnerSta = System.Threading.Thread.CurrentThread.ManagedThreadId });
                button.Execute();
                evidence(new { Phase = "MenuExecuteReturned" });
            }
            finally { foreach (var button in matches) Release(button); Release(bars); Release(editor); }
        }

        private static void FindEmbeddedGitButton(object parent, List<object> matches, int depth)
        {
            if (depth > 6) throw new InvalidOperationException("Bounded CommandBar menu depth exceeded.");
            object controls = null;
            try
            {
                controls = ((dynamic)parent).Controls;
                int count = Convert.ToInt32(((dynamic)controls).Count);
                if (count > 250) throw new InvalidOperationException("Bounded CommandBar control count exceeded.");
                for (int index = 1; index <= count; index++)
                {
                    object control = ((dynamic)controls).Item(index); bool retained = false;
                    try
                    {
                        if (Convert.ToString(((dynamic)control).Tag) == "VBAi.GitHub") { matches.Add(control); retained = true; }
                        else if (Convert.ToInt32(((dynamic)control).Type) == 10) FindEmbeddedGitButton(control, matches, depth + 1);
                    }
                    finally { if (!retained) Release(control); }
                }
            }
            finally { Release(controls); }
        }

        internal void VerifyEmbeddedGitState(EmbeddedGitScope scope)
        {
            Assert.AreEqual(scope.State, ReadEmbeddedState(scope, out _, out _, out _),
                "Embedded capture/checkpoint must preserve exact project identity, source, references, mode, selection and form state.");
        }

        /// <summary>Checks complete imported resources and measured native properties independently of UI status.</summary>
        internal void VerifyEmbeddedImportedForm(EmbeddedGitScope scope, Action<object> evidence)
        {
            RequireEmbeddedProcess(scope);
            WithGitProject(scope.Path, project => {
                var actual = project.Capture();
                evidence(new { Phase = "IndependentPostImportSnapshot", Exact = scope.Baseline.SameAs(actual),
                    Changes = actual.Changes(scope.Baseline), Files = EmbeddedGitSnapshotOracle.Describe(actual) });
                Assert.IsTrue(scope.Baseline.SameAs(actual), "Owner-dispatched import did not preserve the complete snapshot.");
            });
            var layout = ReadGitLayout("EmbeddedForm", scope.Layout);
            var fonts = ReadGitLayoutFonts("EmbeddedForm", scope.Layout);
            evidence(new { Phase = "IndependentPostImportNativeReadback", Properties = layout, Fonts = fonts });
            foreach (var expected in scope.NativeLayout) Assert.AreEqual(expected.Value, layout[expected.Key], expected.Key);
            foreach (var expected in scope.NativeFonts) Assert.AreEqual(expected.Value, fonts[expected.Key], expected.Key);
            Assert.IsFalse(Convert.ToBoolean(((dynamic)workbook).Saved), "An import is an unsaved edit, even after exact checkpoint restoration.");
        }

        private string ReadEmbeddedState(EmbeddedGitScope scope, out Dictionary<string, string> source, out Dictionary<string, int> types, out string referenceRevision)
        {
            object project = null, components = null, references = null, editor = null, pane = null, code = null;
            source = new Dictionary<string, string>(StringComparer.Ordinal); types = new Dictionary<string, int>(StringComparer.Ordinal);
            try
            {
                project = OwnGitProjectRcw(); components = ((dynamic)project).VBComponents;
                foreach (object component in (dynamic)components)
                    try
                    {
                        code = ((dynamic)component).CodeModule; int lines = Convert.ToInt32(((dynamic)code).CountOfLines);
                        string name = Convert.ToString(((dynamic)component).Name);
                        source.Add(name, lines == 0 ? "" : Convert.ToString(((dynamic)code).Lines[1, lines]));
                        types.Add(name, Convert.ToInt32(((dynamic)component).Type));
                    }
                    finally { Release(code); code = null; Release(component); }
                var rows = new List<string>(); references = ((dynamic)project).References;
                foreach (object reference in (dynamic)references)
                    try { rows.Add(Convert.ToString(((dynamic)reference).GUID) + ":" + ((dynamic)reference).Major + ":" + ((dynamic)reference).Minor); Assert.IsFalse((bool)((dynamic)reference).IsBroken); }
                    finally { Release(reference); }
                referenceRevision = string.Join(";", rows.OrderBy(x => x, StringComparer.Ordinal).Select(x => x.ToUpperInvariant()));
                editor = ((dynamic)application).VBE; pane = ((dynamic)editor).ActiveCodePane;
                Assert.AreEqual(2, Convert.ToInt32(((dynamic)project).Mode)); Assert.AreEqual(0, Convert.ToInt32(((dynamic)project).Protection));
                object activeProject = ((dynamic)editor).ActiveVBProject, activeCode = null, activeComponent = null,
                    activeCollection = null, paneProject = null;
                string activeModule;
                try
                {
                    Assert.IsTrue(VbeProjectHostPath.SameProject(project, activeProject));
                    activeCode = ((dynamic)pane).CodeModule; activeComponent = ((dynamic)activeCode).Parent;
                    activeCollection = ((dynamic)activeComponent).Collection; paneProject = ((dynamic)activeCollection).Parent;
                    Assert.IsTrue(VbeProjectHostPath.SameProject(project, paneProject), "The active pane must belong to the same owned project, even when module names match.");
                    activeModule = Convert.ToString(((dynamic)activeComponent).Name);
                    Assert.AreEqual("EmbeddedModule", activeModule);
                }
                finally { Release(paneProject); Release(activeCollection); Release(activeComponent); Release(activeCode); Release(activeProject); }
                int a = 0, b = 0, c = 0, d = 0; ((dynamic)pane).GetSelection(ref a, ref b, ref c, ref d);
                long identity; IntPtr unknown = Marshal.GetIUnknownForObject(project);
                try { identity = unknown.ToInt64(); } finally { Marshal.Release(unknown); }
                return new JavaScriptSerializer().Serialize(new { Identity = identity, Mode = (int)((dynamic)project).Mode,
                    Path = (string)((dynamic)workbook).FullName, Saved = (bool)((dynamic)workbook).Saved,
                    Code = source.OrderBy(x => x.Key).ToArray(), Types = types.OrderBy(x => x.Key).ToArray(), References = referenceRevision,
                    ActiveModule = activeModule, Selection = new[] { a, b, c, d }, Form = ReadGitForm("EmbeddedForm") });
            }
            finally { Release(pane); Release(editor); Release(references); Release(components); Release(project); }
        }
    }
}
