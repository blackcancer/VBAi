using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VBAi;
using VBAi.Tests.Infrastructure;

namespace VBAi.Tests.Integration
{
    [TestClass, TestCategory("Excel")]
    public sealed class MonacoBreakpointExcelTests : EditorUiTestFixture
    {
        [STATestMethod]
        public void InvalidLinesStaySilentAndValidMonacoBreakpointStopsNativeVba()
        {
            using (var host = ExcelVbeFixture.Start())
            using (var fixture = new EditorFixture())
            {
                dynamic excel = UiInvoke.Field<object>(host, "application");
                dynamic book = UiInvoke.Field<object>(host, "workbook");
                dynamic vbe = excel.VBE;
                dynamic project = book.VBProject;
                dynamic component = project.VBComponents.Add(1);
                component.Name = "BreakpointAudit";
                const string source = "Option Explicit\r\nPublic Sub Probe()\r\n    ' Not executable\r\n    Debug.Print 42\r\nEnd Sub";
                component.CodeModule.AddFromString(source);
                string original = (string)component.CodeModule.Lines[1, 5];
                var adapter = new EditorVbeModule((object)vbe, (object)project, (object)component);
                try
                {
                    using (var window = new ModernEditorWindow { Drafts = new EditorDraftStore(fixture.Root) })
                    {
                        var document = MonacoRuntimeTests.Wait(window.OpenModule(adapter));
                        window.Show();
                        MonacoRuntimeTests.Wait(() => window.Ready);
                        MonacoRuntimeTests.Wait(() => MonacoRuntimeTests.Wait(window.Script("snapshots")).Contains(document.Id));
                        UiInvoke.Field<Timer>(window, "timer").Stop();
                        string status = UiInvoke.Field<Label>(window, "status").Text;
                        foreach (int line in new[] { 1, 3 })
                        {
                            Toggle(window, line);
                            StringAssert.Contains(MonacoRuntimeTests.Wait(window.Script("testInfo")), "\"pendingBreakpoints\":0");
                            Assert.AreEqual(status, UiInvoke.Field<Label>(window, "status").Text, "Invalid locations must not display a message.");
                        }
                        Toggle(window, 4);
                        StringAssert.Contains(MonacoRuntimeTests.Wait(window.Script("testInfo")), "\"pendingBreakpoints\":1");
                        adapter.ShowNative(2, 1);
                        vbe.ActiveCodePane.Window.SetFocus();
                        vbe.CommandBars.FindControl(1, 186).Execute();
                        MonacoRuntimeTests.Wait(() => (int)project.Mode == 1);
                        int lineAtBreak = 0, column = 0, end = 0, endColumn = 0;
                        vbe.ActiveCodePane.GetSelection(ref lineAtBreak, ref column, ref end, ref endColumn);
                        Assert.AreEqual(4, lineAtBreak, "Native execution must actually stop on the Monaco breakpoint.");
                        Toggle(window, 4);
                        StringAssert.Contains(MonacoRuntimeTests.Wait(window.Script("testInfo")), "\"pendingBreakpoints\":0");
                        vbe.CommandBars.FindControl(1, 186).Execute();
                        MonacoRuntimeTests.Wait(() => (int)project.Mode == 2);
                        Assert.AreEqual(original, (string)component.CodeModule.Lines[1, 5]);
                    }
                }
                finally
                {
                    try
                    {
                        if ((int)project.Mode == 1)
                            new VbeDebug((object)vbe).ExecuteGlobalDebugCommand(new Request { Project = (string)project.Name, ExpectedMode = 1, Action = "reset" });
                    }
                    finally
                    {
                        foreach (object value in new object[] { component, project, vbe })
                            if (Marshal.IsComObject(value)) Marshal.ReleaseComObject(value);
                    }
                }
            }
        }

        private static void Toggle(ModernEditorWindow window, int line)
        {
            MonacoRuntimeTests.Wait(window.Script("reveal", line, 1));
            var watch = Stopwatch.StartNew();
            MonacoRuntimeTests.Wait(window.Script("command", "vbai.toggle_breakpoint"));
            MonacoRuntimeTests.Wait(() => MonacoRuntimeTests.Wait(window.Script("testInfo")).Contains("\"pendingCommands\":0") &&
                !UiInvoke.Field<bool>(window, "busy"));
            Console.WriteLine("Monaco toggle line " + line + ": " + watch.ElapsedMilliseconds + " ms (single fixture observation).");
        }
    }
}
