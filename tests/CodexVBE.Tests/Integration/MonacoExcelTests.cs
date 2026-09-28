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
                    var refusal = Assert.ThrowsException<InvalidOperationException>(() => guarded.Write(withChange, withChange.Replace("Sub Special", "Sub Renamed")));
                    StringAssert.Contains(refusal.Message, "hidden procedure attributes");
                    Assert.AreEqual(withChange, EditorDocument.Normalize(guarded.Read()));
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
