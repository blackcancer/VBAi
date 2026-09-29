using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Web.Script.Serialization;
using VBAi;
using VBAi.Tests.Infrastructure;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Integration
{
    /// <summary>Qualifie dans Excel visible les références et les fonctions de langage du Monaco réel.</summary>
    [TestClass, TestCategory("MonacoExcel")]
    public sealed class MonacoLanguageExcelTests
    {
        private static readonly JavaScriptSerializer Json = new JavaScriptSerializer();
        private static string Query(ModernEditorWindow window, string operation, string id, int line, int column)
        {
            MonacoRuntimeTests.Wait(window.Browser.CoreWebView2.ExecuteScriptAsync("window.languageProbe=null;window.vbai." + operation + "(" + Json.Serialize(id) + "," + line + "," + column + ").then(value=>window.languageProbe={result:value});"));
            MonacoRuntimeTests.Wait(() => MonacoRuntimeTests.Wait(window.Browser.CoreWebView2.ExecuteScriptAsync("window.languageProbe")) != "null");
            return MonacoRuntimeTests.Wait(window.Browser.CoreWebView2.ExecuteScriptAsync("window.languageProbe"));
        }
        private static string Read(ModernEditorWindow window, string id) => MonacoRuntimeTests.Wait(window.Script("read", id));
        private static void Draft(ModernEditorWindow window, string id, string text)
        {
            var state = (System.Collections.Generic.IDictionary<string, object>)Json.DeserializeObject(Read(window, id));
            Assert.IsTrue(int.Parse(MonacoRuntimeTests.Wait(window.Script("apply", id, Convert.ToInt32(state["version"]), text))) > 0);
        }
        /// <summary>Ajoute puis retire une référence et vérifie blocs, formatage, survol et annulation sans exécuter de macro.</summary>
        [STATestMethod]
        public void LiveReferencesHoverAndAutomaticEditingWorkInTheVisibleExcelRenderer()
        {
            if (Environment.GetEnvironmentVariable("VBAI_EDITOR_LANGUAGE_EXCEL_TEST") != "1") Assert.Inconclusive("Explicit isolated Excel language test opt-in required.");
            if (Process.GetProcessesByName("EXCEL").Length != 0) Assert.Inconclusive("An existing user Excel session must be preserved.");
            dynamic excel = null, workbook = null;
            using (var fixture = new EditorFixture())
            try
            {
                excel = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
                excel.Visible = true; excel.DisplayAlerts = false; excel.EnableEvents = false;
                workbook = excel.Workbooks.Add(); excel.VBE.MainWindow.Visible = true;
                dynamic project = workbook.VBProject;
                dynamic component = project.VBComponents.Add(1); component.Name = "MonacoLanguageFixture";
                component.CodeModule.AddFromString("Option Explicit\r\nPublic Sub Demo()\r\nDim dictionary As Object\r\nDebug.Print 1\r\nEnd Sub");
                string nativeBefore = component.CodeModule.Lines[1, component.CodeModule.CountOfLines];
                using (var window = new ModernEditorWindow { WindowState = System.Windows.Forms.FormWindowState.Maximized })
                {
                    window.Drafts = new EditorDraftStore(fixture.Root);
                    var document = MonacoRuntimeTests.Wait(window.OpenModule(new EditorVbeModule(excel.VBE, (object)project, (object)component)));
                    window.Show(); MonacoRuntimeTests.Wait(() => window.Ready);
                    MonacoRuntimeTests.Wait(() => MonacoRuntimeTests.Wait(window.Script("snapshots")).Contains(document.Id));
                    var synchronizationTimer = UiInvoke.Field<System.Windows.Forms.Timer>(window, "timer");
                    MonacoRuntimeTests.Wait(() => synchronizationTimer.Enabled);
                    synchronizationTimer.Stop();
                    foreach (var scenario in new[] {
                        new[] { "Dim app As Excel.Application", "app.", "Workbooks", "WorksheetFunction" },
                        new[] { "Dim sheet As Excel.Worksheet", "sheet.Range(\"A1\").", "Value2", "Font" },
                        new[] { "Dim sheet As Excel.Worksheet", "sheet.Range(\"A1\").Font.", "Bold", "Size" },
                        new[] { "Dim book As Excel.Workbook", "book.Worksheets.", "Count", "Item" },
                        new[] { "Dim app As Excel.Application", "app.Workbooks(1).Names.", "Add", "Count" },
                        new[] { "Dim value As String", "VBA.Strings.", "Left", "Left$" },
                        new[] { "Dim value As Double", "VBA.Math.", "Abs", "Sqr" } })
                    {
                        Draft(window, document.Id, "Sub Demo()\n" + scenario[0] + "\n" + scenario[1] + "\nEnd Sub");
                        string members = Query(window, "languageInspect", document.Id, 3, scenario[1].Length + 1);
                        foreach (string name in new[] { scenario[2], scenario[3] })
                            Assert.IsTrue(members.Contains("\"Name\":\"" + name + "\""), scenario[1] + " must offer " + name);
                    }
                    Draft(window, document.Id, "Sub Demo()\n \nEnd Sub");
                    string globals = Query(window, "languageInspect", document.Id, 2, 2);
                    foreach (string name in new[] { "MsgBox", "Left", "Abs", "xlUp" })
                        Assert.IsTrue(globals.Contains("\"Name\":\"" + name + "\""), "Global completion must offer " + name);
                    bool hasOffice = false;
                    foreach (dynamic loadedReference in project.References) hasOffice |= (string)loadedReference.Name == "Office";
                    if (!hasOffice)
                    {
                        string[] officePaths = {
                            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonProgramFiles), "Microsoft Shared", "OFFICE16", "MSO.DLL"),
                            Path.Combine(Directory.GetParent((string)excel.Path).FullName, "vfs", "ProgramFilesCommonX64", "Microsoft Shared", "OFFICE16", "MSO.DLL") };
                        string officePath = Array.Find(officePaths, File.Exists);
                        Assert.IsNotNull(officePath, "Installed Office type library must be available for the explicit reference test.");
                        dynamic officeReference = project.References.AddFromFile(officePath);
                        Assert.IsTrue(Query(window, "languageInspect", document.Id, 2, 2).Contains("\"Name\":\"msoTrue\""));
                        project.References.Remove(officeReference);
                        Assert.IsFalse(Query(window, "languageInspect", document.Id, 2, 2).Contains("\"Name\":\"msoTrue\""));
                    }
                    Draft(window, document.Id, "Sub Demo()\nVBA.Strings.Left$\nEnd Sub");
                    StringAssert.Contains(Query(window, "languageHover", document.Id, 2, 17), "Strings.Left$");
                    const string completionText = "Option Explicit\nPublic Sub Demo()\nDim dictionary As Scripting.Dictionary\n    dictionary.\nEnd Sub";
                    Draft(window, document.Id, completionText);
                    string missing = Query(window, "languageInspect", document.Id, 4, 16);
                    Assert.IsFalse(missing.Contains("\"Name\":\"Add\""));
                    dynamic reference = project.References.AddFromFile(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System), "scrrun.dll"));
                    string loaded = Query(window, "languageInspect", document.Id, 4, 16);
                    StringAssert.Contains(loaded, "\"Name\":\"Add\""); StringAssert.Contains(loaded, "Scripting"); StringAssert.Contains(loaded, " As ");
                    project.References.Remove(reference);
                    Assert.IsFalse(Query(window, "languageInspect", document.Id, 4, 16).Contains("\"Name\":\"Add\""));
                    Draft(window, document.Id, "Sub Demo()\nDebug.Print\nEnd Sub");
                    StringAssert.Contains(Query(window, "languageInspect", document.Id, 2, 7), "\"Name\":\"Print\"");
                    StringAssert.Contains(Query(window, "languageHover", document.Id, 2, 10), "Immediate window");
                    reference = project.References.AddFromFile(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System), "scrrun.dll"));
                    Draft(window, document.Id, completionText.Replace("dictionary.\n", "dictionary.Add\n"));
                    string hover = Query(window, "languageHover", document.Id, 4, 18);
                    StringAssert.Contains(hover, "Scripting"); StringAssert.Contains(hover, "scrrun.dll");
                    project.References.Remove(reference);
                    Draft(window, document.Id, "Sub Generated()");
                    MonacoRuntimeTests.Wait(window.Script("reveal", 1, 16)); MonacoRuntimeTests.Wait(window.Script("type", "\n"));
                    MonacoRuntimeTests.Wait(() => Read(window, document.Id).Contains("End Sub"));
                    var state = (System.Collections.Generic.IDictionary<string, object>)Json.DeserializeObject(Read(window, document.Id));
                    Assert.AreEqual("Sub Generated()\n    \nEnd Sub", state["text"]);
                    StringAssert.Contains(MonacoRuntimeTests.Wait(window.Script("position")), "\"column\":5");
                    MonacoRuntimeTests.Wait(window.Script("command", "undo"));
                    MonacoRuntimeTests.Wait(() => ((System.Collections.Generic.IDictionary<string, object>)Json.DeserializeObject(Read(window, document.Id)))["text"].Equals("Sub Generated()"));
                    Draft(window, document.Id, "Sub Demo()\nFor i = 1 To 3\nEnd Sub");
                    MonacoRuntimeTests.Wait(window.Script("reveal", 2, 15)); MonacoRuntimeTests.Wait(window.Script("type", "\n"));
                    MonacoRuntimeTests.Wait(() => Read(window, document.Id).Contains("Next i"));
                    StringAssert.Contains(MonacoRuntimeTests.Wait(window.Script("position")), "\"column\":9");
                    Draft(window, document.Id, "Sub Demo()\nif(value = 1)then\nDebug.Print \"value(then)\" ' keep\nend if\nend sub");
                    MonacoRuntimeTests.Wait(window.Script("command", "editor.action.formatDocument"));
                    MonacoRuntimeTests.Wait(() => Read(window, document.Id).Contains("If (value = 1) Then"));
                    state = (System.Collections.Generic.IDictionary<string, object>)Json.DeserializeObject(Read(window, document.Id));
                    StringAssert.Contains((string)state["text"], "        Debug.Print \"value(then)\" ' keep");
                    Assert.AreEqual(nativeBefore, (string)component.CodeModule.Lines[1, component.CodeModule.CountOfLines], "Language services must not execute or rewrite native VBA during this inspection.");
                    window.Close(); MonacoRuntimeTests.Wait(() => window.IsDisposed);
                }
            }
            finally
            {
                if ((object)workbook != null) { workbook.Close(false); Marshal.FinalReleaseComObject((object)workbook); }
                if ((object)excel != null) { excel.Quit(); Marshal.FinalReleaseComObject((object)excel); }
            }
        }
    }
}
