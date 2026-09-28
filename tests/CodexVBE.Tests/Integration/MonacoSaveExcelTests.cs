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
    public sealed class MonacoSaveExcelTests
    {
        [DllImport("user32.dll")] private static extern bool SetForegroundWindow(IntPtr window);
        [STATestMethod]
        public void SaveSynchronizesAndPersistsTargetWorkbookAndPreservesCodeAfterCancellationOrFailure()
        {
            if (Environment.GetEnvironmentVariable("VBAI_EDITOR_EXCEL_TEST") != "1") Assert.Inconclusive("Explicit disposable Excel opt-in required.");
            if (Process.GetProcessesByName("EXCEL").Length != 0) Assert.Inconclusive("Close existing Excel processes before this isolated test.");
            dynamic excel = null, book = null, other = null;
            using (var fixture = new EditorFixture())
            try
            {
                Directory.CreateDirectory(fixture.Root);
                string path = Path.Combine(fixture.Root, "SaveTarget.xlsm");
                excel = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
                excel.Visible = true; excel.DisplayAlerts = false; excel.EnableEvents = false;
                book = excel.Workbooks.Add();
                dynamic component = book.VBProject.VBComponents.Add(1); component.Name = "SaveFixture";
                component.CodeModule.AddFromString(fixture.Code.Replace("\n", "\r\n"));
                var adapter = new EditorVbeModule(excel.VBE, book.VBProject, component);
                using (var window = new ModernEditorWindow())
                {
                    var persistence = new VbeProjectComponents(adapter.Vbe, null,
                        new ScopedExcelHostProbe { Application = (object)excel, ProcessId = adapter.HostProcessId });
                    Func<EditorVbeModule, bool?> hostSaved = module =>
                    {
                        dynamic state = persistence.PersistenceStatus(module.ProjectName);
                        Assert.IsTrue((bool)state.HostAvailable);
                        Assert.AreEqual((bool)book.Saved, (bool)state.HostSaved);
                        return (bool?)state.HostSaved;
                    };
                    window.NativeHostSaved = hostSaved;
                    window.Drafts = new EditorDraftStore(fixture.Root);
                    var doc = MonacoRuntimeTests.Wait(window.OpenModule(adapter)); window.Show();
                    MonacoRuntimeTests.Wait(() => window.Ready);
                    UiInvoke.Field<System.Windows.Forms.Timer>(window, "timer").Stop();
                    MonacoRuntimeTests.Wait(window.Script("reveal", 3, 1));
                    MonacoRuntimeTests.Wait(window.Script("insert", "    ' save fixture\n"));
                    var nativeSave = window.NativeSave;
                    // Returning without changing Saved/path models dismissal of the native dialog.
                    window.NativeSave = module => { };
                    Assert.ThrowsException<InvalidOperationException>(() => MonacoRuntimeTests.Wait(window.SaveDocument(doc.Id)));
                    StringAssert.Contains(doc.Text, "save fixture");
                    StringAssert.Contains(adapter.Read(), "save fixture");
                    window.NativeSave = module => { throw new IOException("simulated host failure"); };
                    Assert.ThrowsException<IOException>(() => MonacoRuntimeTests.Wait(window.SaveDocument(doc.Id)));
                    StringAssert.Contains(doc.Text, "save fixture");
                    book.SaveAs(path, 52);
                    // Project persistence alone must never stand in for whole-document persistence.
                    Assert.IsTrue((bool)book.VBProject.Saved);
                    window.NativeSave = module => { };
                    window.NativeHostSaved = module => false;
                    var hostFailure = Assert.ThrowsException<InvalidOperationException>(() => MonacoRuntimeTests.Wait(window.SaveDocument(doc.Id)));
                    StringAssert.Contains(hostFailure.Message, "host document was not saved");
                    StringAssert.Contains(doc.Text, "save fixture");
                    window.NativeHostSaved = module => null;
                    MonacoRuntimeTests.Wait(window.SaveDocument(doc.Id));
                    Assert.AreEqual(UiText.Get("The native Save command finished, but the host document's saved state could not be verified."),
                        UiInvoke.Field<System.Windows.Forms.Label>(window, "status").Text);
                    window.NativeHostSaved = hostSaved;
                    other = excel.Workbooks.Add();
                    other.Activate();
                    int saves = 0;
                    window.NativeSave = module => { saves++; nativeSave(module); };
                    MonacoRuntimeTests.Wait(window.Script("reveal", 3, 1));
                    MonacoRuntimeTests.Wait(window.Script("insert", "    ' persisted by native save\n"));
                    window.Activate(); window.Browser.Focus(); SetForegroundWindow(window.Handle);
                    MonacoRuntimeTests.Wait(() => window.ContainsFocus);
                    MonacoRuntimeTests.Wait(window.Script("command", "vbai.save"));
                    try { MonacoRuntimeTests.Wait(() => !UiInvoke.Field<bool>(window, "busy") &&
                        adapter.Read().Contains("persisted by native save") && (bool)book.Saved); }
                    catch (AssertFailedException)
                    {
                        Assert.Fail("Save diagnostic: calls=" + saves + ", hostSaved=" + (bool)book.Saved +
                            ", projectSaved=" + (bool)book.VBProject.Saved + ", dirty=" + doc.Dirty +
                            ", status=" + UiInvoke.Field<System.Windows.Forms.Label>(window, "status").Text);
                    }
                    Assert.IsTrue((bool)book.Saved);
                    Assert.AreEqual(1, saves, "The Save action must route exactly once through Monaco synchronization and the native Save command.");
                    Assert.AreEqual(0, ((string)other.Path).Length, "Save must not save the unrelated active workbook.");
                }
                book.Close(false); Marshal.FinalReleaseComObject((object)book); book = null;
                book = excel.Workbooks.Open(path);
                string persisted = book.VBProject.VBComponents.Item("SaveFixture").CodeModule.Lines[1, 8];
                StringAssert.Contains(persisted, "persisted by native save");
            }
            finally
            {
                if (book != null) { try { book.Close(false); } catch (COMException) { } Marshal.FinalReleaseComObject((object)book); }
                if (other != null) { try { other.Close(false); } catch (COMException) { } Marshal.FinalReleaseComObject((object)other); }
                if (excel != null) { try { excel.Quit(); } catch (COMException) { } Marshal.FinalReleaseComObject((object)excel); }
                GC.Collect(); GC.WaitForPendingFinalizers();
            }
        }
    }
}
