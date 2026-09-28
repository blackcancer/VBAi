using System;
using System.Collections;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace CodexVBE
{
    internal sealed partial class ModernEditorWindow
    {
        private int lastDebugMode = -1;
        private string lastDebugDocument;
        private async Task ObserveDebugMode()
        {
            if (busy || Current == null || !(Current.Module is EditorVbeModule native)) return;
            int mode = (int)((dynamic)native.Project).Mode;
            if (lastDebugMode == mode && lastDebugDocument == Current.Id) return;
            lastDebugMode = mode; lastDebugDocument = Current.Id;
            if (mode != 1)
            { foreach (var doc in documents.Values.ToArray()) await Script("execution", doc.Id, 0); return; }
            if (!Current.Dirty && !Current.Conflict)
                await EditorCommand(new EditorMessage { id = Current.Id, version = versions[Current.Id], name = "show_next_statement" });
        }
        private async Task EditorCommand(EditorMessage message)
        {
            if (busy || !documents.TryGetValue(message.id ?? "", out var document) || !(document.Module is EditorVbeModule native)) return;
            await CaptureDocuments();
            if (versions[document.Id] != message.version) throw new InvalidOperationException("The editor changed before the command. Retry at the current location.");
            await ProcessDocuments(true);
            if (document.Dirty || document.Conflict) throw new InvalidOperationException("Synchronize or resolve the draft before compiling or debugging.");
            if (message.name != "compile" && versions[document.Id] != message.version) throw new InvalidOperationException("VBA reformatted the source. Retry the command at its current position.");
            busy = true;
            try
            {
                await Task.Yield(); // Native commands must never execute inside a WebView callback.
                var debugger = new VbeDebug(native.Vbe);
                int mode = (int)((dynamic)native.Project).Mode;
                if (message.name == "compile")
                {
                    if (mode != 2) throw new InvalidOperationException("Compilation requires design mode.");
                    native.ShowNative(1, 1);
                    VbeDebugWindows.EnsureNoCompileDialog();
                    string diagnostic;
                    using (var completed = new ManualResetEventSlim())
                    {
                        var observe = Task.Run(() => VbeDebugWindows.AwaitCompileDialog(completed));
                        Exception failure = null;
                        try { debugger.CompileProject(new Request { Project = native.ProjectName, ExpectedMode = 2 }); }
                        catch (Exception error) { failure = error; }
                        finally { completed.Set(); }
                        diagnostic = await observe;
                        if (failure != null) throw failure;
                    }
                    foreach (var item in documents.Values) await Script("diagnostics", item.Id, versions[item.Id], Array.Empty<object>());
                    if (!string.IsNullOrWhiteSpace(diagnostic))
                    {
                        dynamic pane = ((dynamic)native.Vbe).ActiveCodePane;
                        var target = native.Sibling((string)pane.CodeModule.Parent.Name);
                        if (!target.IsComponent((object)pane.CodeModule.Parent)) throw new InvalidOperationException("The compiler selection belongs to a different project.");
                        int line = 1, column = 1, endLine = 1, endColumn = 1;
                        pane.GetSelection(ref line, ref column, ref endLine, ref endColumn);
                        var doc = await OpenModule(target);
                        await Script("diagnostics", doc.Id, versions[doc.Id], new[] { new { message = diagnostic, startLineNumber = line, startColumn = column, endLineNumber = endLine, endColumn = Math.Max(column + 1, endColumn), severity = 8, source = "VBA compiler" } });
                        await Script("reveal", line, column);
                    }
                    else status.Text = UiText.Get("Compilation finished: no native diagnostics observed. Macros were not executed.");
                }
                else if (message.name == "show_next_statement")
                {
                    debugger.ExecuteGlobalDebugCommand(new Request { Project = native.ProjectName, ExpectedMode = mode, Action = message.name });
                    await Task.Yield();
                    dynamic pane = ((dynamic)native.Vbe).ActiveCodePane;
                    int line = 1, column = 1, endLine = 1, endColumn = 1;
                    pane.GetSelection(ref line, ref column, ref endLine, ref endColumn);
                    var target = native.Sibling((string)pane.CodeModule.Parent.Name);
                    if (!target.IsComponent((object)pane.CodeModule.Parent)) throw new InvalidOperationException("The execution selection belongs to a different project.");
                    var doc = await OpenModule(target);
                    foreach (var item in documents.Values) await Script("execution", item.Id, item == doc ? line : 0);
                }
                else
                {
                    if (!new[] { "toggle_breakpoint", "step_into", "step_over", "step_out" }.Contains(message.name)) throw new ArgumentException("Unknown editor command.");
                    object control = null;
                    // Search the native command inventory, never emit global keyboard shortcuts.
                    for (int offset = 0; offset < 2000 && control == null; offset += 200)
                    {
                        var page = ((IEnumerable)debugger.ListCommands(null, offset, 200)).Cast<object>().ToArray();
                        foreach (dynamic item in page)
                            if ((bool)item.Enabled && VbeDebug.IsAllowed(message.name, (string)item.Caption, mode)) { control = item; break; }
                        if (page.Length < 200) break;
                    }
                    if (control == null) throw new InvalidOperationException("The native debug command is unavailable in the current mode.");
                    string source = native.Read();
                    debugger.InvokeCommand(new Request { Project = native.ProjectName, Module = native.ModuleName, ExpectedMode = mode,
                        ExpectedSha256 = EditorDocument.Hash(source), StartLine = Math.Max(1, message.line), Action = message.name, ControlId = (int)((dynamic)control).Id, ControlCaption = (string)((dynamic)control).Caption });
                    // VBIDE cannot enumerate breakpoints. This is deliberately a hollow request marker.
                    if (message.name == "toggle_breakpoint") await Script("breakpointRequested", document.Id, message.line);
                    else { lastDebugMode = -1; foreach (var item in documents.Values) await Script("execution", item.Id, 0); }
                }
                Activate();
            }
            finally { busy = false; }
        }
    }
}
