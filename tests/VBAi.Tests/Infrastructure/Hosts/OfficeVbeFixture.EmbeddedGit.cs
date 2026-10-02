using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Integration
{
    internal sealed partial class OfficeVbeFixture
    {
        /// <summary>Prepares a saved inert Word project and independent bridge exports, without external Git capture.</summary>
        internal ExcelVbeFixture.EmbeddedGitScope PrepareWordEmbeddedGit(string marker, Action<object> record)
        {
            Assert.AreEqual("Word", Kind);
            RequireUsableOwnedHost(); RequireOwnedDocument();
            PrepareGitMarker(DocumentPath, marker);
            NativeExecutionUnsettled = true;
            SaveNative();
            NativeExecutionUnsettled = false;
            var scope = new ExcelVbeFixture.EmbeddedGitScope {
                Path = DocumentPath, Marker = marker, Cache = MacroGitRepository.ScopeDirectory(DocumentPath),
                Code = new Dictionary<string, string>(StringComparer.Ordinal), Types = new Dictionary<string, int>(StringComparer.Ordinal)
            };
            Assert.IsFalse(Directory.Exists(scope.Cache), "A fresh disposable Word scope must not reuse a cache.");
            var components = new List<VbaGitComponent>();
            var files = new Dictionary<string, byte[]>(StringComparer.Ordinal);
            string exports = Path.Combine(Root, "independent-bridge-baseline"); Directory.CreateDirectory(exports);
            foreach (var item in Items("list_modules", "Project", DocumentPath).OrderBy(row => Convert.ToString(row["Name"]), StringComparer.Ordinal))
            {
                string name = Convert.ToString(item["Name"]); int type = Convert.ToInt32(item["Type"]);
                Assert.IsTrue(type == 1 || type == 2 || type == 100, "Q024 does not include UserForm resources.");
                var source = Data("read_module", "Project", DocumentPath, "Module", name);
                scope.Code.Add(name, Convert.ToString(source["Code"])); scope.Types.Add(name, type);
                var component = new VbaGitComponent { Name = name, Type = type };
                string text;
                if (type == 100) text = scope.Code[name];
                else
                {
                    string path = Path.Combine(exports, component.FileName);
                    var state = Data("component_properties", "Project", DocumentPath, "Module", name);
                    Data("export_component", "Project", DocumentPath, "Module", name, "Path", path, "ExpectedComponentVersion", state["Version"]);
                    var codec = Data("inspect_code_file", "Path", path);
                    text = Encoding.GetEncoding(Convert.ToInt32(codec["SystemAnsiCodePage"]), EncoderFallback.ExceptionFallback,
                        DecoderFallback.ExceptionFallback).GetString(File.ReadAllBytes(path));
                    record(new { Phase = "IndependentBridgeExport", Path = path, Codec = codec });
                }
                files.Add(component.FileName, VbaGitSnapshot.Utf8.GetBytes(text.Replace("\r\n", "\n").Replace("\r", "\n")));
                components.Add(component);
            }
            object project = null, references = null, editor = null, window = null;
            try
            {
                project = ((dynamic)document).VBProject; references = ((dynamic)project).References;
                var ids = new List<string>();
                foreach (object reference in (dynamic)references)
                    try
                    {
                        Assert.IsFalse((bool)((dynamic)reference).IsBroken);
                        ids.Add(Convert.ToString(((dynamic)reference).GUID).ToUpperInvariant() + ":" +
                            Convert.ToInt32(((dynamic)reference).Major) + ":" + Convert.ToInt32(((dynamic)reference).Minor));
                    }
                    finally { Release(reference); }
                scope.References = string.Join(";", ids.OrderBy(id => id, StringComparer.Ordinal));
                editor = ((dynamic)application).VBE; window = ((dynamic)editor).MainWindow;
                scope.VbeHandle = new IntPtr(Convert.ToInt64(((dynamic)window).HWnd));
                uint pid; scope.ThreadId = GetWindowThreadProcessId(scope.VbeHandle, out pid);
                Assert.AreEqual((uint)ProcessId, pid); Assert.AreNotEqual(0u, scope.ThreadId);
            }
            finally { Release(window); Release(editor); Release(references); Release(project); }
            scope.Baseline = new VbaGitSnapshot(new VbaGitManifest { Components = components.ToArray(), References = scope.References }, files);
            var markerCode = Data("read_module", "Project", DocumentPath, "Module", "QualificationMarker");
            Data("select_code", "Project", DocumentPath, "Module", "QualificationMarker", "StartLine", 1, "ExpectedSha256", markerCode["Sha256"]);
            Assert.AreEqual(DocumentPath, Data("debug_state", "Project", DocumentPath)["SelectedHostPath"]);
            record(new { Phase = "WordScopePrepared", ProcessId, scope.ThreadId, VbeHandle = scope.VbeHandle.ToInt64(),
                scope.Path, scope.Cache, Files = EmbeddedGitSnapshotOracle.Describe(scope.Baseline),
                MacroExecutions = 0, ExternalGitCaptures = 0 });
            return scope;
        }

        /// <summary>Checks only the retained process and its frozen native window, safe for the MTA UI worker.</summary>
        internal void RequireWordEmbeddedOwner(ExcelVbeFixture.EmbeddedGitScope scope)
        {
            if (!owned || ownedProcess == null || ownedProcess.HasExited)
                throw new InvalidOperationException("The original owned Word process is no longer available.");
            uint pid; uint tid = GetWindowThreadProcessId(scope.VbeHandle, out pid);
            EmbeddedGitUiProtocol.RequireOwner(ProcessId, scope.ThreadId, scope.VbeHandle.ToInt64(), (int)pid, tid, scope.VbeHandle.ToInt64());
        }

        /// <summary>Invokes exactly one tagged production Git menu; no keyboard, focus or pointer automation.</summary>
        internal void ExecuteWordGitMenu(Action<object> record)
        {
            object editor = null, bars = null; var buttons = new List<object>();
            try
            {
                editor = ((dynamic)application).VBE; bars = ((dynamic)editor).CommandBars;
                foreach (object bar in (dynamic)bars)
                    try { if (Convert.ToInt32(((dynamic)bar).Type) == 1) FindWordGitButton(bar, buttons, 0); }
                    finally { Release(bar); }
                Assert.AreEqual(1, buttons.Count, "The exact Word VBAi.GitHub button must be unique.");
                dynamic button = buttons[0]; Assert.IsTrue((bool)button.Enabled);
                record(new { Phase = "MenuExecuteIntent", Tag = (string)button.Tag, ProcessId });
                button.Execute();
                record(new { Phase = "MenuExecuteReturned", ProcessId });
            }
            finally { foreach (object button in buttons) Release(button); Release(bars); Release(editor); }
        }

        private void FindWordGitButton(object parent, List<object> buttons, int depth)
        {
            if (depth > 6) throw new InvalidOperationException("Bounded Word menu depth exceeded.");
            object controls = null;
            try
            {
                controls = ((dynamic)parent).Controls; int count = Convert.ToInt32(((dynamic)controls).Count);
                if (count > 250) throw new InvalidOperationException("Bounded Word menu count exceeded.");
                for (int index = 1; index <= count; index++)
                {
                    object control = ((dynamic)controls).Item(index); bool retained = false;
                    try
                    {
                        if (Convert.ToString(((dynamic)control).Tag) == "VBAi.GitHub") { buttons.Add(control); retained = true; }
                        else if (Convert.ToInt32(((dynamic)control).Type) == 10) FindWordGitButton(control, buttons, depth + 1);
                    }
                    finally { if (!retained) Release(control); }
                }
            }
            finally { Release(controls); }
        }

        internal void VerifyWordEmbeddedSource(ExcelVbeFixture.EmbeddedGitScope scope)
        {
            RequireWordEmbeddedOwner(scope); RequireOwnedDocument();
            var rows = Items("list_modules", "Project", scope.Path);
            CollectionAssert.AreEquivalent(scope.Code.Keys.ToArray(), rows.Select(row => Convert.ToString(row["Name"])).ToArray());
            foreach (var row in rows)
            {
                string name = Convert.ToString(row["Name"]);
                Assert.AreEqual(scope.Types[name], Convert.ToInt32(row["Type"]));
                Assert.AreEqual(scope.Code[name], Data("read_module", "Project", scope.Path, "Module", name)["Code"]);
            }
            Assert.AreEqual(scope.Path, Data("debug_state", "Project", scope.Path)["SelectedHostPath"]);
            Assert.IsTrue((bool)((dynamic)document).Saved, "Capture/checkpoint must not dirty the saved Word document.");
            object project = null, references = null;
            try
            {
                project = ((dynamic)document).VBProject;
                Assert.AreEqual(2, Convert.ToInt32(((dynamic)project).Mode));
                Assert.AreEqual(0, Convert.ToInt32(((dynamic)project).Protection));
                references = ((dynamic)project).References; var ids = new List<string>();
                foreach (object reference in (dynamic)references)
                    try
                    {
                        Assert.IsFalse((bool)((dynamic)reference).IsBroken);
                        ids.Add(Convert.ToString(((dynamic)reference).GUID).ToUpperInvariant() + ":" +
                            Convert.ToInt32(((dynamic)reference).Major) + ":" + Convert.ToInt32(((dynamic)reference).Minor));
                    }
                    finally { Release(reference); }
                Assert.AreEqual(scope.References, string.Join(";", ids.OrderBy(id => id, StringComparer.Ordinal)));
            }
            finally { Release(references); Release(project); }
        }
    }
}
