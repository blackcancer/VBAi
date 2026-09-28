using System;
using CodexVBE;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace CodexVBE.Tests.Unit.Editor
{
    [TestClass, TestCategory("Unit")]
    public sealed class EditorAttributeGuardTests
    {
        [TestMethod]
        public void ModuleVariableAttributesAreNotDiscardedByReplacement()
        {
            const string source = "Public Value As Long";
            Assert.ThrowsException<InvalidOperationException>(() => EditorVbeModule.EnsureAttributeDeclarationsUntouched("Attribute Value.VB_VarHelpID = 1", source, Tuple.Create(1, 1, "Public Other As Long")));
        }
        [TestMethod]
        public void BodyAndAdjacentInsertionsPreserveAttributedDeclarations()
        {
            string source = "Public Sub Special( _\n ByVal value As Long)\n Debug.Print value\nEnd Sub";
            string export = "Attribute Special.VB_Description = \"Keep\"";
            EditorVbeModule.EnsureAttributeDeclarationsUntouched(export, source, EditorDocument.Difference(source, source.Replace("Print value", "Print 42")));
            EditorVbeModule.EnsureAttributeDeclarationsUntouched(export, source, EditorDocument.Difference(source, "' before\n" + source));
            Assert.ThrowsException<InvalidOperationException>(() => EditorVbeModule.EnsureAttributeDeclarationsUntouched(export, source, EditorDocument.Difference(source, source.Replace("Special", "Other"))));
            Assert.ThrowsException<InvalidOperationException>(() => EditorVbeModule.EnsureAttributeDeclarationsUntouched(export, source, Tuple.Create(2, 0, "' split")));
        }
    }
}

namespace CodexVBE.Tests.Unit.Editor
{
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Runtime.InteropServices;
    using System.Text;
    using System.Threading.Tasks;
    using System.Windows.Forms;
    using CodexVBE.Tests.Infrastructure;

    [TestClass]
    public sealed class EditorVbeModuleTests
    {
        private const string Single = "Public Sub Hello()\nAttribute Hello.VB_Description = \"Keep\"\n    Debug.Print 1\nEnd Sub";
        private const string Multi = "Public Sub Hello( _\n    Optional ByVal count As Long = 1)\nAttribute Hello.VB_Description = \"Keep\"\n    Debug.Print count\nEnd Sub";
        private static string Read(EditorVbeContract f) => EditorDocument.Normalize(f.Adapter.Read());
        private static string Changed(string code) => code.Replace("count As Long = 1", "count As Long = 2");

        [STATestMethod]
        public void ComponentIdentityUsesIUnknownForDistinctNativeWrappers()
        {
            var same = typeof(EditorVbeModule).GetMethod("Same", System.Reflection.BindingFlags.Static | System.Reflection.BindingFlags.NonPublic);
            object first = Activator.CreateInstance(Type.GetTypeFromProgID("StdFont", true));
            object second = Activator.CreateInstance(Type.GetTypeFromProgID("StdFont", true));
            object alias = null; IntPtr pointer = IntPtr.Zero;
            try
            {
                pointer = Marshal.GetIUnknownForObject(first);
                alias = Marshal.GetUniqueObjectForIUnknown(pointer);
                Assert.AreNotSame(first, alias);
                Assert.IsTrue((bool)same.Invoke(null, new[] { first, alias }));
                Assert.IsFalse((bool)same.Invoke(null, new[] { first, second }));
                Assert.IsFalse((bool)same.Invoke(null, new[] { first, new object() }));
                Assert.IsFalse((bool)same.Invoke(null, new[] { new object(), first }));
            }
            finally
            {
                if (pointer != IntPtr.Zero) Marshal.Release(pointer);
                if (alias != null) Marshal.FinalReleaseComObject(alias);
                Marshal.FinalReleaseComObject(first); Marshal.FinalReleaseComObject(second);
            }
        }

        [STATestMethod]
        public async Task IdentitySourcesNeighborsPaneAndProcessFollowTheNativeObjects()
        {
            var f = new EditorVbeContract(); var adapter = f.Adapter;
            Assert.AreEqual("Project1 · Module1", adapter.Name); Assert.AreSame(f.Project, adapter.Project); Assert.AreSame(f.Vbe, adapter.Vbe);
            Assert.AreSame(f.Original, adapter.Component); Assert.IsTrue(adapter.IsComponent(f.Original)); Assert.IsFalse(adapter.IsComponent(new object()));
            string sessionKey = adapter.Key; Assert.IsTrue(sessionKey.EndsWith("|Module1"));
            foreach (var path in new[] { (string)null, " ", "relative.xlsm", "C:/Temp/Test.xlsm" })
            {
                f.Project.Path = path; string key = adapter.Key;
                if (path == "C:/Temp/Test.xlsm") Assert.IsTrue(key.StartsWith(Path.GetFullPath(path).ToUpperInvariant())); else Assert.AreEqual(sessionKey, key);
                Assert.AreEqual(string.IsNullOrEmpty(path) ? "Project1" : path, adapter.ProjectName);
            }
            f.Project.FailFileName = true; Assert.AreEqual(sessionKey, adapter.Key); Assert.AreEqual("Project1", adapter.ProjectName); f.Project.FailFileName = false;
            var empty = new EditorVbeContract.Component { Name = "Empty", Type = 2, Collection = f.Project.VBComponents }; f.Project.VBComponents.Items.Add(empty);
            var sources = await adapter.Sources(); Assert.AreEqual(2, sources.Length); Assert.AreEqual("", sources[1].Text);
            Assert.AreSame(empty, adapter.Sibling("eMpTy").Component); Assert.ThrowsException<InvalidOperationException>(() => adapter.Sibling("Gone"));
            adapter.CloseNativeWindow(); adapter.EnsureNativeWindow(); Assert.IsTrue(f.Original.CodeModule.CodePane.Window.Visible);
            adapter.CloseNativeWindow(); Assert.AreEqual(1, f.Original.CodeModule.CodePane.Window.Closes);
            adapter.EnsureNativeWindow(); f.Original.CodeModule.CodePane.Window.FailClose = true; adapter.CloseNativeWindow();
            adapter.ShowNative(-10, -2); Assert.AreEqual(1, f.Original.CodeModule.CodePane.Line); Assert.AreEqual(1, f.Original.CodeModule.CodePane.Column);
            adapter.ShowNative(999, 20); Assert.AreEqual(f.Original.CodeModule.CountOfLines, f.Original.CodeModule.CodePane.Line);
            var emptyAdapter = adapter.Sibling("Empty"); emptyAdapter.ShowNative(2, 2); Assert.AreEqual(1, empty.CodeModule.CodePane.Line); Assert.AreEqual("", emptyAdapter.Read());
            Assert.ThrowsException<InvalidOperationException>(() => { var id = adapter.HostProcessId; });
            f.Vbe.MainWindow.HWnd = 1; Assert.ThrowsException<InvalidOperationException>(() => { var id = adapter.HostProcessId; });
            using (var owner = new Form()) { f.Vbe.MainWindow.HWnd = owner.Handle.ToInt32(); Assert.AreEqual(System.Diagnostics.Process.GetCurrentProcess().Id, adapter.HostProcessId); }
            f.Project.VBComponents.Items.Remove(f.Original); Assert.ThrowsException<InvalidOperationException>(() => adapter.Read());
            f.Project.VBComponents.Items.Add(f.Original); f.Vbe.VBProjects.Clear(); Assert.ThrowsException<InvalidOperationException>(() => adapter.Read());
        }

        [TestMethod]
        public void WriteGuardsPlansAndEveryPlainPatchShapePreserveExpectedSource()
        {
            foreach (var mode in new[] { 0, 1, 2 }) foreach (var protection in new[] { 0, 1 })
            {
                var f = new EditorVbeContract(); f.Project.Mode = mode; f.Project.Protection = protection;
                Assert.AreEqual((mode == 1 || mode == 2) && protection == 0, f.Adapter.CanWrite);
                if (!f.Adapter.CanWrite) Assert.ThrowsException<InvalidOperationException>(() => f.Adapter.Write(Read(f), "new"));
            }
            foreach (var shape in new[] { "same", "replace", "insert", "delete", "many", "empty" })
            {
                var f = new EditorVbeContract(); string before = Read(f), after = shape == "same" ? before : shape == "replace" ? before.Replace("Print 1", "Print 2") :
                    shape == "insert" ? "' header\n" + before : shape == "delete" ? before.Replace("    Debug.Print 1\n", "") : shape == "empty" ? "" : "' first\n' second";
                Assert.AreEqual(after, EditorDocument.Normalize(f.Adapter.WritePrepared(before, after, new EditorSyncPlan(before, after))));
            }
            var guarded = new EditorVbeContract(); string expected = Read(guarded);
            Assert.ThrowsException<InvalidOperationException>(() => guarded.Adapter.Write("stale", "next"));
            foreach (var plan in new[] { new EditorSyncPlan("wrong", "next"), new EditorSyncPlan(expected, "wrong") })
                Assert.ThrowsException<InvalidOperationException>(() => guarded.Adapter.WritePrepared(expected, "next", plan));
            if (Encoding.Default.CodePage != 65001) Assert.ThrowsException<EncoderFallbackException>(() => guarded.Adapter.Write(expected, expected + "\n' 😀"));
        }

        [TestMethod]
        public void BreakModeAcceptsOnlyOneBodyLineAndKeepsStructuralEditsPending()
        {
            foreach (string shape in new[] { "body", "insert", "multiline", "delete", "declaration" })
            {
                var f = new EditorVbeContract(); string before = Read(f);
                f.Project.Mode = 1;
                string after = shape == "body" ? before.Replace("Print 1", "Print 2") :
                    shape == "insert" ? "' header\n" + before :
                    shape == "multiline" ? before.Replace("Print 1", "Print 2\n    Debug.Print 3") :
                    shape == "delete" ? before.Replace("    Debug.Print 1\n", "") : before.Replace("Hello()", "Changed()");
                if (shape == "body")
                {
                    Assert.AreEqual(after, EditorDocument.Normalize(f.Adapter.Write(before, after)));
                    Assert.AreSame(f.Original, f.Adapter.Component);
                    Assert.AreEqual(1, f.Project.Mode);
                }
                else
                {
                    Assert.ThrowsException<InvalidOperationException>(() => f.Adapter.Write(before, after), shape);
                    Assert.AreEqual(before, Read(f));
                    Assert.AreSame(f.Original, f.Adapter.Component);
                }
            }
        }

        [TestMethod]
        public void PlainMutationFailuresRestoreSingleAndMultipleLinesOrReportBothFailures()
        {
            foreach (var shape in new[] { "replace", "insert", "delete", "many" })
            {
                var f = new EditorVbeContract(); string before = Read(f), after = shape == "replace" ? before.Replace("Print 1", "Print 2") :
                    shape == "insert" ? before + "\n' inserted" : shape == "delete" ? before.Replace("    Debug.Print 1\n", "") : "' altered\n' again";
                int calls = 0; f.Original.CodeModule.AfterOperation = op => { if (calls++ == 0) throw new IOException("After mutation"); };
                var failure = Assert.ThrowsException<InvalidOperationException>(() => f.Adapter.Write(before, after));
                StringAssert.Contains(failure.Message, "original source was restored"); Assert.AreEqual(before, Read(f));
            }
            foreach (var restoration in new[] { "throw", "mismatch", "noop" })
            {
                var f = new EditorVbeContract(); string before = Read(f); int calls = 0;
                f.Original.CodeModule.AfterOperation = op => { calls++; if (calls == 1 || restoration == "throw") throw new IOException("operation"); if (restoration == "mismatch") f.Original.CodeModule.Raw += "\n' mismatch"; };
                if (restoration == "noop") f.Original.CodeModule.BeforeOperation = op => { throw new IOException("Before mutation"); };
                var failure = Assert.ThrowsException<InvalidOperationException>(() => f.Adapter.Write(before, before.Replace("Print 1", "Print 2")));
                if (restoration == "noop") Assert.AreEqual(before, Read(f)); else Assert.IsInstanceOfType(failure.InnerException, typeof(AggregateException));
            }
        }

        [TestMethod]
        public void SingleLineAttributesReloadAndReadBackWithoutReplacingTheComponent()
        {
            var f = new EditorVbeContract(1, Single); string before = Read(f), after = before.Replace("Hello()", "Hello(ByVal value As Long)");
            Assert.AreEqual(after, EditorDocument.Normalize(f.Adapter.Write(before, after))); Assert.AreSame(f.Original, f.Adapter.Component);
            StringAssert.Contains(f.Original.CodeModule.Raw, "Attribute Hello.VB_Description");
            var body = new EditorVbeContract(1, Single); string bodyBefore = Read(body);
            Assert.AreEqual(bodyBefore.Replace("Print 1", "Print 2"), EditorDocument.Normalize(body.Adapter.Write(bodyBefore, bodyBefore.Replace("Print 1", "Print 2"))));
        }

        [TestMethod]
        public void AttributedLoadFailuresRestoreMetadataOrRetainARecoveryExport()
        {
            foreach (var kind in new[] { "load", "source", "metadata", "rollback", "rollback-source", "rollback-metadata" })
            {
                var f = new EditorVbeContract(1, Single); string before = Read(f); int adds = 0;
                f.Original.CodeModule.AfterOperation = op => {
                    if (op != "add") return; adds++;
                    if (adds == 1 && (kind == "load" || kind == "rollback")) throw new IOException("Load failed");
                    if (adds == 1 && (kind == "source" || kind == "rollback-source")) f.Original.CodeModule.Raw += "\n' mismatch";
                    if (adds == 1 && (kind == "metadata" || kind == "rollback-metadata")) f.Original.CodeModule.Raw = f.Original.CodeModule.Raw.Replace("Keep", "Lost");
                    if (adds > 1 && kind == "rollback") throw new IOException("Rollback failed");
                    if (adds > 1 && kind == "rollback-source") f.Original.CodeModule.Raw += "\n' mismatch";
                    if (adds > 1 && kind == "rollback-metadata") f.Original.CodeModule.Raw = f.Original.CodeModule.Raw.Replace("Keep", "Lost");
                };
                var failure = Assert.ThrowsException<InvalidOperationException>(() => f.Adapter.Write(before, before.Replace("Hello()", "Hello(ByVal value As Long)")));
                if (kind.StartsWith("rollback")) { Assert.IsInstanceOfType(failure.InnerException, typeof(AggregateException)); CleanupRecovery(failure); }
                else { Assert.AreEqual(before, Read(f)); StringAssert.Contains(f.Original.CodeModule.Raw, "Keep"); }
            }
            foreach (var mode in new[] { true, false })
            {
                var f = new EditorVbeContract(1, Single); string before = Read(f); int exports = 0;
                f.Original.AfterExport = path => { if (++exports == 1) { if (mode) f.Project.Mode = 1; else f.Original.CodeModule.Raw += "\n' external"; } };
                Assert.ThrowsException<InvalidOperationException>(() => f.Adapter.Write(before, before.Replace("Hello()", "Hello(ByVal value As Long)")));
            }
        }

        [TestMethod]
        public void MultilineReplacementRetainsCodeAttributesAndDesignerAcrossComponentTypes()
        {
            foreach (int type in new[] { 1, 2, 3 })
            {
                var f = new EditorVbeContract(type, Multi); string before = Read(f), after = Changed(before);
                Assert.AreEqual(after, EditorDocument.Normalize(f.Adapter.Write(before, after))); Assert.AreNotSame(f.Original, f.Adapter.Component);
                Assert.AreEqual(1, f.Project.VBComponents.Items.Count); Assert.AreEqual("Module1", f.Adapter.ModuleName);
                StringAssert.Contains(((EditorVbeContract.Component)f.Adapter.Component).CodeModule.Raw, "Attribute Hello.VB_Description");
            }
            var document = new EditorVbeContract(100, Multi); string original = Read(document);
            Assert.ThrowsException<InvalidOperationException>(() => document.Adapter.Write(original, Changed(original))); Assert.AreEqual(original, Read(document));
        }

        [TestMethod]
        public void StagedValidationRefusesDivergentTypeCodeAttributesAndDesignerWithoutRemovingOriginal()
        {
            foreach (var fault in new[] { "type", "code", "attributes", "designer", "empty", "separator-only" })
            {
                var f = new EditorVbeContract(fault == "designer" || fault == "separator-only" ? 3 : 1, Multi); string before = Read(f);
                f.Project.VBComponents.AfterImport = (candidate, path) => { if (!Path.GetFileName(path).StartsWith("candidate")) return;
                    if (fault == "type") candidate.Type = 99; if (fault == "code") candidate.CodeModule.Raw += "\n' wrong";
                    if (fault == "attributes") candidate.CodeModule.Raw = candidate.CodeModule.Raw.Replace("Keep", "Lost");
                    if (fault == "designer") candidate.Designer.Caption = "Different"; if (fault == "empty") candidate.CodeModule.Raw = ""; };
                if (fault == "separator-only") f.Project.VBComponents.AfterImport = (candidate, path) => { if (Path.GetFileName(path).StartsWith("candidate")) candidate.CodeModule.Raw = "\n"; };
                Assert.ThrowsException<InvalidOperationException>(() => f.Adapter.Write(before, Changed(before)));
                Assert.AreSame(f.Original, f.Adapter.Component); Assert.AreEqual(before, Read(f)); Assert.AreEqual(1, f.Project.VBComponents.Items.Count);
            }
        }

        [TestMethod]
        public void PreparedAndRetiredChecksPreserveConcurrentSourceModeNameAndDesigner()
        {
            foreach (var fault in new[] { "mode-prepared", "code-prepared", "name-prepared", "mode-retired", "code-retired", "name-retired", "candidate-name", "designer-retired", "candidate-code", "candidate-attributes" })
            {
                var f = new EditorVbeContract(fault == "designer-retired" ? 3 : 1, Multi); string before = Read(f);
                f.Adapter.AttributeRewriteCheckpoint = phase => { if (phase != "prepared") return;
                    if (fault == "mode-prepared") f.Project.Mode = 1; if (fault == "code-prepared") f.Original.CodeModule.Raw += "\n' concurrent";
                    if (fault == "name-prepared") f.Original.Name = "External"; };
                f.Project.VBComponents.AfterImport = (candidate, path) => candidate.AfterName = name => { if (name != "Module1") return;
                    if (fault == "mode-retired") f.Project.Mode = 1; if (fault == "code-retired") f.Original.CodeModule.Raw += "\n' concurrent";
                    if (fault == "name-retired") f.Original.Name = "External";
                    if (fault == "candidate-name") { candidate.AfterName = null; candidate.Name = "ChangedName"; }
                    if (fault == "designer-retired") f.Original.Designer.Caption = "External caption";
                    if (fault == "candidate-code") candidate.CodeModule.Raw += "\n' wrong";
                    if (fault == "candidate-attributes") candidate.CodeModule.Raw = candidate.CodeModule.Raw.Replace("Keep", "Lost"); };
                Assert.ThrowsException<InvalidOperationException>(() => f.Adapter.Write(before, Changed(before)));
                Assert.AreSame(f.Original, f.Adapter.Component); Assert.AreEqual(1, f.Project.VBComponents.Items.Count);
                Assert.AreEqual(fault.StartsWith("name-") ? "External" : "Module1", f.Original.Name);
                if (fault.StartsWith("code-")) StringAssert.Contains(Read(f), "concurrent"); else Assert.AreEqual(before, Read(f));
                if (fault == "designer-retired") Assert.AreEqual("External caption", f.Original.Designer.Caption);
            }
        }

        [TestMethod]
        public void RemovalAndImportFailuresRestoreOriginalOrRetainItsRecoverableBackup()
        {
            foreach (var fault in new[] { "import-before", "import-after", "remove-before", "remove-after", "removed-checkpoint", "open", "candidate-remove", "restore-import", "restore-code" })
            {
                var f = new EditorVbeContract(1, Multi); string before = Read(f); int removes = 0;
                f.Project.VBComponents.ImportThrowsBefore = fault == "import-before"; f.Project.VBComponents.ImportThrowsAfter = fault == "import-after";
                f.Project.VBComponents.BeforeRemove = component => { if (fault == "remove-before" && ReferenceEquals(component, f.Original)) throw new IOException("Before remove");
                    if (fault == "candidate-remove" && !ReferenceEquals(component, f.Original)) throw new IOException("Candidate cleanup"); };
                f.Project.VBComponents.AfterRemove = component => { removes++; if (fault == "remove-after" && ReferenceEquals(component, f.Original)) throw new IOException("After remove"); };
                f.Project.VBComponents.AfterImport = (candidate, path) => { if (Path.GetFileName(path).StartsWith("candidate")) {
                    if (fault == "open") candidate.CodeModule.CodePane.Window.OnVisible = () => { throw new IOException("Window reopen"); };
                    if (fault == "candidate-remove") candidate.Type = 99;
                } else if (fault == "restore-code") candidate.CodeModule.Raw += "\n' wrong"; };
                f.Adapter.AttributeRewriteCheckpoint = phase => { if (phase == "removed" && (fault == "removed-checkpoint" || fault.StartsWith("restore-"))) {
                    if (fault == "restore-import") f.Project.VBComponents.ImportThrowsBefore = true; throw new IOException("After removal"); } };
                var failure = Assert.ThrowsException<Exception>(() => { try { f.Adapter.Write(before, Changed(before)); } catch (Exception error) { throw new Exception("Captured", error); } }).InnerException;
                if (fault == "candidate-remove" || fault.StartsWith("restore-")) { Assert.IsInstanceOfType(failure, typeof(AggregateException)); CleanupRecovery(failure); }
                else { Assert.AreEqual(before, Read(f)); Assert.AreEqual(1, f.Project.VBComponents.Items.Count); }
            }
        }

        [TestMethod]
        public void AttributeGuardsCoverQualifiedIncompleteGlobalAndAdjacentDeclarations()
        {
            foreach (var source in new[] { "Public Private", "Public Sub", "Property Get", "Option Explicit", "Public Function Other()\nEnd Function", "Public Value As Long" })
                EditorVbeModule.EnsureAttributeDeclarationsUntouched("Attribute Missing.VB_Description = 1", source, Tuple.Create(20, 0, "' after"));
            foreach (var source in new[] { "Public Static Sub Hello()\nEnd Sub", "Private Function Hello() As Long\nEnd Function", "Friend Property Get Hello() As Long\nEnd Property" })
            {
                Assert.ThrowsException<InvalidOperationException>(() => EditorVbeModule.EnsureAttributeDeclarationsUntouched("Attribute Hello.VB_Description = 1", source, Tuple.Create(1, 1, "changed")));
                EditorVbeModule.EnsureAttributeDeclarationsUntouched("Attribute Hello.VB_Description = 1", source, Tuple.Create(1, 0, "' before"));
                EditorVbeModule.EnsureAttributeDeclarationsUntouched("Attribute Hello.VB_Description = 1", source, Tuple.Create(20, 1, "' after"));
            }
            EditorVbeModule.EnsureAttributeDeclarationsUntouched("", "Public Value As Long", Tuple.Create(1, 1, "changed"));
            Assert.ThrowsException<InvalidOperationException>(() => EditorVbeModule.EnsureAttributeDeclarationsUntouched("Attribute Value.VB_Description = 1", "Public Value As _\n Long", Tuple.Create(2, 0, "' split")));
            foreach (var edit in new[] { Tuple.Create(1, 1, "' before"), Tuple.Create(5, 1, "' after"), Tuple.Create(2, 0, "' adjacent before"), Tuple.Create(5, 0, "' adjacent after") })
                EditorVbeModule.EnsureAttributeDeclarationsUntouched("Attribute Value.VB_Description = 1", "' before\nPublic Value As _\n Long\n' after", edit);
        }

        private static void CleanupRecovery(Exception failure)
        {
            string text = failure.ToString(); var match = System.Text.RegularExpressions.Regex.Match(text, @"(?:retained at |recovery export: )(.+?\.(?:bas|cls|frm))(?=\s|$)");
            Assert.IsTrue(match.Success, text); string path = match.Groups[1].Value.Trim();
            if (path.EndsWith(" --->")) path = path.Substring(0, path.Length - 5);
            Assert.IsTrue(File.Exists(path), path); string directory = Path.GetFullPath(Path.GetDirectoryName(path));
            string temp = Path.GetFullPath(Path.GetTempPath()); Assert.IsTrue(directory.StartsWith(temp, StringComparison.OrdinalIgnoreCase));
            Assert.IsTrue(Path.GetFileName(directory).StartsWith("VBAi-attributes-", StringComparison.Ordinal)); Directory.Delete(directory, true);
        }
    }
}
