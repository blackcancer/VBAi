using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using CodexVBE;
using CodexVBE.Tests.Infrastructure;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace CodexVBE.Tests.Integration
{
    [TestClass, TestCategory("MonacoExcel")]
    public sealed class MonacoExcelTests
    {
        [DllImport("user32.dll")] private static extern bool SetForegroundWindow(IntPtr window);
        [DllImport("user32.dll")] private static extern IntPtr GetForegroundWindow();
        [STATestMethod]
        public void DisposableExcelModuleRoundTripsRealMonacoChangesAndDetectsConcurrentNativeEdits()
        {
            if (Environment.GetEnvironmentVariable("VBAI_EDITOR_EXCEL_TEST") != "1") Assert.Inconclusive("Explicit disposable Excel opt-in required.");
            if (Process.GetProcessesByName("EXCEL").Length != 0) Assert.Inconclusive("Close existing Excel processes before this isolated test.");
            dynamic excel = null, workbook = null;
            using (var fixture = new EditorFixture())
            try
            {
                excel = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
                excel.Visible = true; excel.DisplayAlerts = false; excel.EnableEvents = false;
                excel.VBE.MainWindow.Visible = true;
                workbook = excel.Workbooks.Add();
                object project = workbook.VBProject;
                dynamic component = ((dynamic)project).VBComponents.Add(1); component.Name = "MonacoFixture";
                component.CodeModule.AddFromString(fixture.Code.Replace("\n", "\r\n"));
                var adapter = new EditorVbeModule(excel.VBE, project, component);
                using (var window = new ModernEditorWindow())
                {
                    window.Drafts = new EditorDraftStore(fixture.Root);
                    // This test controls synchronization boundaries rather than relying on timer timing.
                    var doc = MonacoRuntimeTests.Wait(window.OpenModule(adapter)); window.Show();
                    MonacoRuntimeTests.Wait(() => window.Ready); UiInvoke.Field<System.Windows.Forms.Timer>(window, "timer").Stop();
                    MonacoRuntimeTests.Wait(() => MonacoRuntimeTests.Wait(window.Script("snapshots")).Contains(doc.Id));
                    UiInvoke.Field<System.Windows.Forms.Timer>(window, "timer").Stop();
                    // A real compiler diagnostic must reach Monaco and then disappear after correction.
                    component.CodeModule.ReplaceLine(3, "    UnknownDiagnosticVariable = 1");
                    MonacoRuntimeTests.Wait(window.ProcessDocuments(false));
                    MonacoRuntimeTests.Wait(window.Script("command", "vbai.compile"));
                    MonacoRuntimeTests.Wait(() => MonacoRuntimeTests.Wait(window.Script("testInfo")).Contains("\"markers\":1"));
                    MonacoRuntimeTests.Wait(() => !UiInvoke.Field<bool>(window, "busy"));
                    int diagnosticLine = 0, diagnosticColumn = 0, diagnosticEnd = 0, diagnosticEndColumn = 0;
                    excel.VBE.ActiveCodePane.GetSelection(ref diagnosticLine, ref diagnosticColumn, ref diagnosticEnd, ref diagnosticEndColumn);
                    Assert.AreEqual(3, diagnosticLine, "The native compiler must select the invalid identifier.");
                    component.CodeModule.ReplaceLine(3, "    Debug.Print 1");
                    MonacoRuntimeTests.Wait(window.ProcessDocuments(false));
                    StringAssert.Contains(MonacoRuntimeTests.Wait(window.Script("testInfo")), "\"markers\":0");
                    // Exercise the real Monaco action -> WebView message -> VBE compilation path.
                    MonacoRuntimeTests.Wait(window.Script("command", "vbai.compile"));
                    MonacoRuntimeTests.Wait(() => UiInvoke.Field<System.Windows.Forms.Label>(window, "status").Text == UiText.Get("Compilation finished: no native diagnostics observed. Macros were not executed."));
                    Assert.AreEqual(2, (int)((dynamic)project).Mode, "Compilation must leave the project in design mode.");
                    string assistantCommand = null; ChatAttachment assistantAttachment = null;
                    window.AssistantAction += (command, attachment) => { assistantCommand = command; assistantAttachment = attachment; };
                    foreach (string action in new[] { "expliquer", "corriger", "refactoriser" })
                    {
                        assistantCommand = null;
                        MonacoRuntimeTests.Wait(window.Script("command", "vbai." + action));
                        MonacoRuntimeTests.Wait(() => assistantCommand != null);
                        Assert.AreEqual("/" + action, assistantCommand);
                        Assert.AreEqual(doc.Id, assistantAttachment.EditorDocumentId);
                        Assert.AreEqual(adapter.ProjectName, assistantAttachment.Project);
                        StringAssert.Contains(assistantAttachment.Text, "Debug.Print 1");
                    }
                    // A Monaco toggle must stop actual VBA execution on the requested line.
                    MonacoRuntimeTests.Wait(window.Script("reveal", 3, 1));
                    MonacoRuntimeTests.Wait(window.Script("command", "vbai.toggle_breakpoint"));
                    MonacoRuntimeTests.Wait(() => MonacoRuntimeTests.Wait(window.Script("testInfo")).Contains("\"pendingBreakpoints\":1"));
                    MonacoRuntimeTests.Wait(() => !UiInvoke.Field<bool>(window, "busy"));
                    adapter.ShowNative(2, 1);
                    excel.VBE.ActiveCodePane.Window.SetFocus();
                    excel.VBE.CommandBars.FindControl(1, 186).Execute();
                    MonacoRuntimeTests.Wait(() => (int)((dynamic)project).Mode == 1);
                    int breakLine = 0, breakColumn = 0, breakEnd = 0, breakEndColumn = 0;
                    excel.VBE.ActiveCodePane.GetSelection(ref breakLine, ref breakColumn, ref breakEnd, ref breakEndColumn);
                    Assert.AreEqual(3, breakLine, "Execution must hit the breakpoint requested from Monaco.");
                    // Invoke the same native action again without keyboard or pointer simulation.
                    MonacoRuntimeTests.Wait(window.Script("command", "vbai.toggle_breakpoint"));
                    MonacoRuntimeTests.Wait(() => MonacoRuntimeTests.Wait(window.Script("testInfo")).Contains("\"pendingBreakpoints\":0"));
                    MonacoRuntimeTests.Wait(() => !UiInvoke.Field<bool>(window, "busy"));
                    new VbeDebug(excel.VBE).ExecuteGlobalDebugCommand(new Request { Project = adapter.ProjectName, ExpectedMode = 1, Action = "reset" });
                    MonacoRuntimeTests.Wait(() => (int)((dynamic)project).Mode == 2);
                    // Enter only our disposable Debug.Print fixture, then drive the Monaco debug actions.
                    adapter.ShowNative(3, 1);
                    dynamic step = null;
                    foreach (dynamic command in (System.Collections.IEnumerable)new VbeDebug(excel.VBE).ListCommands(null, 0, 200))
                        if ((bool)command.Enabled && VbeDebug.IsAllowed("step_into", (string)command.Caption, 1)) { step = command; break; }
                    Assert.IsNotNull((object)step, "Native Step Into command must be present.");
                    excel.VBE.CommandBars.FindControl(1, (int)step.Id).Execute();
                    MonacoRuntimeTests.Wait(() => (int)((dynamic)project).Mode == 1);
                    MonacoRuntimeTests.Wait(window.Script("command", "vbai.show_next_statement"));
                    MonacoRuntimeTests.Wait(() => MonacoRuntimeTests.Wait(window.Script("testInfo")).Contains("\"executionMarkers\":1"));
                    MonacoRuntimeTests.Wait(() => !UiInvoke.Field<bool>(window, "busy"));
                    // A simple correction is propagated while paused, without resetting execution.
                    var versions = UiInvoke.Field<System.Collections.Generic.Dictionary<string, int>>(window, "versions");
                    MonacoRuntimeTests.Wait(window.Script("apply", doc.Id, versions[doc.Id], doc.Text.Replace("Debug.Print 1", "Debug.Print 2")));
                    MonacoRuntimeTests.Wait(window.ProcessDocuments(true));
                    StringAssert.Contains(adapter.Read(), "Debug.Print 2");
                    Assert.IsFalse(doc.Dirty);
                    Assert.AreEqual(1, (int)((dynamic)project).Mode);
                    Func<int> nativeLine = () => { int a = 0, b = 0, c = 0, d = 0; excel.VBE.ActiveCodePane.GetSelection(ref a, ref b, ref c, ref d); return a; };
                    foreach (string action in new[] { "vbai.step_into", "vbai.step_over" })
                    {
                        int stoppedLine = nativeLine();
                        MonacoRuntimeTests.Wait(window.Script("command", action));
                        MonacoRuntimeTests.Wait(() => !UiInvoke.Field<bool>(window, "busy") && nativeLine() != stoppedLine);
                    }
                    MonacoRuntimeTests.Wait(window.Script("command", "vbai.step_out"));
                    MonacoRuntimeTests.Wait(() => (int)((dynamic)project).Mode == 2);
                    MonacoRuntimeTests.Wait(() => !UiInvoke.Field<bool>(window, "busy"));
                    MonacoRuntimeTests.Wait(window.Script("reveal", 3, 1));
                    MonacoRuntimeTests.Wait(window.Script("insert", "    ' accented é\n"));
                    MonacoRuntimeTests.Wait(() => doc.Text.Contains("accented é"));
                    MonacoRuntimeTests.Wait(window.ProcessDocuments(true));
                    StringAssert.Contains(adapter.Read(), "accented é"); Assert.IsFalse(doc.Dirty);
                    Directory.CreateDirectory(fixture.Root);
                    string saved = Path.Combine(fixture.Root, "MonacoRoundtrip.xlsm");
                    workbook.SaveAs(saved, 52);
                    string attributed = Path.Combine(fixture.Root, "Attributed.bas");
                    File.WriteAllText(attributed, "Attribute VB_Name = \"Attributed\"\r\nPublic Sub Special()\r\nAttribute Special.VB_Description = \"Preserve me\"\r\nEnd Sub\r\n", System.Text.Encoding.Default);
                    dynamic special = ((dynamic)project).VBComponents.Import(attributed);
                    var guarded = new EditorVbeModule(excel.VBE, project, special);
                    string originalSpecial = EditorDocument.Normalize(guarded.Read());
                    guarded.Write(originalSpecial, originalSpecial + "\n' change");
                    string withChange = EditorDocument.Normalize(guarded.Read());
                    string verifyExport = Path.Combine(fixture.Root, "attribute-verification.bas");
                    special.Export(verifyExport);
                    StringAssert.Contains(File.ReadAllText(verifyExport, System.Text.Encoding.Default), "Attribute Special.VB_Description = \"Preserve me\"");
                    guarded.Write(withChange, withChange.Replace("Sub Special", "Sub Renamed"));
                    Assert.IsTrue(guarded.IsComponent((object)special), "Component COM identity must survive attributed edits.");
                    File.Delete(verifyExport); special.Export(verifyExport);
                    StringAssert.Contains(File.ReadAllText(verifyExport, System.Text.Encoding.Default), "Attribute Renamed.VB_Description = \"Preserve me\"");
                    StringAssert.Contains(guarded.Read(), "Sub Renamed");
                    string renamed = EditorDocument.Normalize(guarded.Read());
                    guarded.Write(renamed, renamed.Replace("Sub Renamed", "Sub Special"));
                    File.Delete(verifyExport); special.Export(verifyExport);
                    StringAssert.Contains(File.ReadAllText(verifyExport, System.Text.Encoding.Default), "Attribute Special.VB_Description = \"Preserve me\"");
                    // Multiline declarations and changed signatures retain opaque procedure metadata.
                    string simple = EditorDocument.Normalize(guarded.Read());
                    string complex = simple.Replace("Public Sub Special()", "Public Sub Special( _\n    Optional ByVal count As Long = 1)");
                    var attributedDocument = MonacoRuntimeTests.Wait(window.OpenModule(guarded));
                    MonacoRuntimeTests.Wait(() => MonacoRuntimeTests.Wait(window.Script("snapshots")).Contains(attributedDocument.Id));
                    object oldSpecial = special;
                    guarded.Write(simple, complex); special = guarded.Component;
                    Assert.IsFalse(guarded.IsComponent(oldSpecial), "Multiline attributes require a staged replacement.");
                    MonacoRuntimeTests.Wait(window.ProcessDocuments(false));
                    Assert.AreSame(attributedDocument, MonacoRuntimeTests.Wait(window.OpenModule(new EditorVbeModule(excel.VBE, project, special))));
                    StringAssert.Contains(attributedDocument.Text, "Optional ByVal count As Long = 1");
                    UiInvoke.Call(typeof(ModernEditorWindow), "CloseModuleClick", window, window, EventArgs.Empty);
                    MonacoRuntimeTests.Wait(() => !System.Linq.Enumerable.Any(window.Documents, item => item.Id == attributedDocument.Id));
                    MonacoRuntimeTests.Wait(window.OpenModule(adapter));
                    File.Delete(verifyExport); special.Export(verifyExport);
                    StringAssert.Contains(File.ReadAllText(verifyExport, System.Text.Encoding.Default), "Attribute Special.VB_Description = \"Preserve me\"");
                    Assert.IsTrue(guarded.IsComponent((object)special));
                    // The same code-only reload must preserve a class component's identity and attributes.
                    string classPath = Path.Combine(fixture.Root, "AttributedClass.cls");
                    File.WriteAllText(classPath, "VERSION 1.0 CLASS\r\nBEGIN\r\n  MultiUse = -1\r\nEND\r\nAttribute VB_Name = \"AttributedClass\"\r\nAttribute VB_PredeclaredId = True\r\nPublic Property Get Value() As Long\r\nAttribute Value.VB_UserMemId = 0\r\n    Value = 42\r\nEnd Property\r\n", System.Text.Encoding.Default);
                    dynamic classComponent = ((dynamic)project).VBComponents.Import(classPath);
                    var classAdapter = new EditorVbeModule(excel.VBE, project, classComponent);
                    string classBefore = EditorDocument.Normalize(classAdapter.Read());
                    classAdapter.Write(classBefore, classBefore.Replace("Value() As Long", "Value() As Double"));
                    Assert.IsTrue(classAdapter.IsComponent((object)classComponent));
                    string classVerify = Path.Combine(fixture.Root, "class-verified.cls"); classComponent.Export(classVerify);
                    string classAfter = File.ReadAllText(classVerify, System.Text.Encoding.Default);
                    StringAssert.Contains(classAfter, "Attribute VB_PredeclaredId = True");
                    StringAssert.Contains(classAfter, "Attribute Value.VB_UserMemId = 0");
                    string classSingle = EditorDocument.Normalize(classAdapter.Read());
                    string classMulti = classSingle.Replace("Value() As Double", "Value( _\n    Optional ByVal index As Long = 0) As Double");
                    classAdapter.Write(classSingle, classMulti); classComponent = classAdapter.Component;
                    File.Delete(classVerify); classComponent.Export(classVerify);
                    StringAssert.Contains(File.ReadAllText(classVerify, System.Text.Encoding.Default), "Attribute Value.VB_UserMemId = 0");
                    // Refusal must preserve a user's concurrent rename, not restore the old name.
                    string renameBefore = EditorDocument.Normalize(guarded.Read());
                    string oldName = (string)((dynamic)guarded.Component).Name;
                    object retainedComponent = guarded.Component;
                    int renameComponentCount = ((dynamic)project).VBComponents.Count;
                    guarded.AttributeRewriteCheckpoint = phase => { if (phase == "prepared") ((dynamic)retainedComponent).Name = "ConcurrentRenameProbe"; };
                    try
                    {
                        Assert.ThrowsException<InvalidOperationException>(() => guarded.Write(renameBefore,
                            renameBefore.Replace("Optional ByVal count As Long = 1", "Optional ByVal count As Long = 2")));
                        Assert.IsTrue(guarded.IsComponent(retainedComponent));
                        Assert.AreEqual("ConcurrentRenameProbe", (string)((dynamic)retainedComponent).Name);
                        Assert.AreEqual(renameBefore, EditorDocument.Normalize(guarded.Read()));
                        Assert.AreEqual(renameComponentCount, (int)((dynamic)project).VBComponents.Count);
                    }
                    finally
                    {
                        guarded.AttributeRewriteCheckpoint = null;
                        ((dynamic)retainedComponent).Name = oldName;
                    }
                    // Force failure after removal to qualify recovery using the actual Excel importer.
                    string recoveryBefore = EditorDocument.Normalize(guarded.Read());
                    string recoveryAfter = recoveryBefore.Replace("Optional ByVal count As Long = 1", "Optional ByVal count As Long = 2");
                    int componentCount = ((dynamic)project).VBComponents.Count;
                    guarded.AttributeRewriteCheckpoint = phase => { if (phase == "removed") throw new InvalidOperationException("Injected post-removal failure"); };
                    Assert.ThrowsException<InvalidOperationException>(() => guarded.Write(recoveryBefore, recoveryAfter));
                    guarded.AttributeRewriteCheckpoint = null; special = guarded.Component;
                    Assert.AreEqual(recoveryBefore, EditorDocument.Normalize(guarded.Read()));
                    Assert.AreEqual(componentCount, (int)((dynamic)project).VBComponents.Count);
                    File.Delete(verifyExport); special.Export(verifyExport);
                    StringAssert.Contains(File.ReadAllText(verifyExport, System.Text.Encoding.Default), "Attribute Special.VB_Description = \"Preserve me\"");
                    // Host document identity and UserForm designer controls survive code-only rewrites.
                    dynamic sheetComponent = ((dynamic)project).VBComponents.Item((string)workbook.Worksheets.Item(1).CodeName);
                    dynamic formComponent = ((dynamic)project).VBComponents.Add(3);
                    dynamic button = formComponent.Designer.Controls.Add("Forms.CommandButton.1", "PreservedButton"); button.Caption = "Preserved designer";
                    foreach (object item in new object[] { sheetComponent, formComponent })
                    {
                        dynamic nativeComponent = item;
                        string fixtureCode = Path.Combine(fixture.Root, "inplace-" + nativeComponent.Name + ".bas");
                        File.WriteAllText(fixtureCode, "Public Sub MetadataProbe()\r\nAttribute MetadataProbe.VB_Description = \"Keep host identity\"\r\nEnd Sub\r\n", System.Text.Encoding.Default);
                        nativeComponent.CodeModule.AddFromFile(fixtureCode);
                        var inplace = new EditorVbeModule(excel.VBE, project, item);
                        string oldText = EditorDocument.Normalize(inplace.Read());
                        inplace.Write(oldText, oldText.Replace("MetadataProbe()", "MetadataProbe(Optional ByVal count As Long = 1)"));
                        Assert.IsTrue(inplace.IsComponent(item));
                        string verified = Path.Combine(fixture.Root, "inplace-" + nativeComponent.Name + ".export"); nativeComponent.Export(verified);
                        StringAssert.Contains(File.ReadAllText(verified, System.Text.Encoding.Default), "Attribute MetadataProbe.VB_Description = \"Keep host identity\"");
                    }
                    Assert.AreEqual("Preserved designer", (string)formComponent.Designer.Controls.Item("PreservedButton").Caption);
                    Assert.AreEqual((string)sheetComponent.Name, (string)workbook.Worksheets.Item(1).CodeName);
                    var formAdapter = new EditorVbeModule(excel.VBE, project, formComponent);
                    string formBefore = EditorDocument.Normalize(formAdapter.Read());
                    string formAfter = formBefore.Replace("MetadataProbe(Optional ByVal count As Long = 1)", "MetadataProbe( _\n    Optional ByVal count As Long = 1)");
                    formAdapter.Write(formBefore, formAfter); formComponent = formAdapter.Component;
                    Assert.AreEqual("Preserved designer", (string)formComponent.Designer.Controls.Item("PreservedButton").Caption);
                    string concurrentBefore = EditorDocument.Normalize(formAdapter.Read());
                    object unchangedForm = formAdapter.Component;
                    formAdapter.AttributeRewriteCheckpoint = phase => { if (phase == "prepared") ((dynamic)unchangedForm).Designer.Controls.Item("PreservedButton").Caption = "Concurrent Designer edit"; };
                    Assert.ThrowsException<InvalidOperationException>(() => formAdapter.Write(concurrentBefore, concurrentBefore.Replace("count As Long = 1", "count As Long = 2")));
                    formAdapter.AttributeRewriteCheckpoint = null;
                    Assert.IsTrue(formAdapter.IsComponent(unchangedForm));
                    Assert.AreEqual("Concurrent Designer edit", (string)((dynamic)unchangedForm).Designer.Controls.Item("PreservedButton").Caption);
                    ((dynamic)unchangedForm).Designer.Controls.Item("PreservedButton").Caption = "Preserved designer";
                    string formRecoveryBefore = EditorDocument.Normalize(formAdapter.Read());
                    int formComponentCount = ((dynamic)project).VBComponents.Count;
                    formAdapter.AttributeRewriteCheckpoint = phase => { if (phase == "removed") throw new InvalidOperationException("Injected UserForm rollback"); };
                    Assert.ThrowsException<InvalidOperationException>(() => formAdapter.Write(formRecoveryBefore, formRecoveryBefore.Replace("count As Long = 1", "count As Long = 2")));
                    formAdapter.AttributeRewriteCheckpoint = null; formComponent = formAdapter.Component;
                    Assert.AreEqual(formRecoveryBefore, EditorDocument.Normalize(formAdapter.Read()));
                    Assert.AreEqual(formComponentCount, (int)((dynamic)project).VBComponents.Count);
                    Assert.AreEqual("Preserved designer", (string)formComponent.Designer.Controls.Item("PreservedButton").Caption);
                    var sheetAdapter = new EditorVbeModule(excel.VBE, project, sheetComponent);
                    string sheetBefore = EditorDocument.Normalize(sheetAdapter.Read());
                    Assert.ThrowsException<InvalidOperationException>(() => sheetAdapter.Write(sheetBefore, sheetBefore.Replace("MetadataProbe(Optional ByVal count As Long = 1)", "MetadataProbe( _\n    Optional ByVal count As Long = 1)")));
                    Assert.AreEqual(sheetBefore, EditorDocument.Normalize(sheetAdapter.Read()));
                    Assert.IsTrue(sheetAdapter.IsComponent((object)sheetComponent));
                    ((dynamic)project).VBComponents.Remove(formComponent);
                    ((dynamic)project).VBComponents.Remove(classComponent);
                    ((dynamic)project).VBComponents.Remove(special);
                    MonacoRuntimeTests.Wait(window.Script("reveal", 3, 1));
                    MonacoRuntimeTests.Wait(window.Script("insert", "    ' local pending\n"));
                    MonacoRuntimeTests.Wait(() => doc.Dirty);
                    component.CodeModule.ReplaceLine(1, "Option Explicit ' changed in VBA");
                    MonacoRuntimeTests.Wait(window.ProcessDocuments(true));
                    Assert.IsTrue(doc.Conflict); StringAssert.Contains(adapter.Read(), "changed in VBA");
                    Assert.IsFalse(adapter.Read().Contains("local pending")); StringAssert.Contains(doc.Text, "local pending");
                    // COM object identity survives a rename and refuses a removed/replaced component.
                    component.Name = "MonacoRenamed"; StringAssert.Contains(adapter.Name, "MonacoRenamed");
                    ((dynamic)project).VBComponents.Remove(component);
                    Assert.ThrowsException<InvalidOperationException>(() => adapter.Read());
                    window.Close(); MonacoRuntimeTests.Wait(() => window.IsDisposed);
                    Assert.IsNotNull(window.Drafts.Recover(doc.RecoveryKey), "Removed module must still retain its dirty draft.");
                    workbook.Close(false); Marshal.FinalReleaseComObject((object)workbook); workbook = null;
                    workbook = excel.Workbooks.Open(saved);
                    string persisted = workbook.VBProject.VBComponents.Item("MonacoFixture").CodeModule.Lines[1, 5];
                    StringAssert.Contains(persisted, "accented é");
                    Assert.IsFalse(persisted.Contains("local pending"));
                }
            }
            catch (Exception error) { Console.WriteLine("PRIMARY EXCEL FAILURE: " + error); throw; }
            finally
            {
                if ((object)workbook != null) { try { workbook.Close(false); } catch (Exception cleanup) { Console.WriteLine("Workbook cleanup: " + cleanup.Message); } Marshal.FinalReleaseComObject((object)workbook); }
                if ((object)excel != null) { try { excel.Quit(); } catch (Exception cleanup) { Console.WriteLine("Excel cleanup: " + cleanup.Message); } Marshal.FinalReleaseComObject((object)excel); }
                GC.Collect(); GC.WaitForPendingFinalizers(); GC.Collect(); GC.WaitForPendingFinalizers();
            }
        }
    }
}
