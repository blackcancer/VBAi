using System;
using System.Collections;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace CodexVBE
{
    /// <summary>Suit le mode d’exécution VBE et exécute les commandes de compilation ou de débogage native.</summary>
internal sealed partial class ModernEditorWindow
    {
        /// <summary>Dernier mode VBE observé pour le document actif.</summary>
private int lastDebugMode = -1;
                /// <summary>Checks that compilation starts without an existing native diagnostic.</summary>
        internal Action<int> EnsureCompileDialogAbsent = VbeDebugWindows.EnsureNoCompileDialog;
        /// <summary>Observes the host diagnostic until its compilation command completes.</summary>
        internal Func<ManualResetEventSlim, int, string> ObserveCompileDialog = VbeDebugWindows.AwaitCompileDialog;
        /// <summary>Identifiant du document associé au dernier mode observé.</summary>
private string lastDebugDocument, lastDebugPosition;
        /// <summary>Stores the last execution line,last execution version used by ModernEditorWindow.</summary>
private int lastExecutionLine, lastExecutionVersion;
        /// <summary>Performs the debug position operation for ModernEditorWindow.</summary>
/// <param name="native">The native used by this operation.</param>
/// <returns>The result produced by this operation.</returns>
private string DebugPosition(EditorVbeModule native)
        {
            dynamic pane = ((dynamic)native.Vbe).ActiveCodePane;
            if (pane == null) return "";
            int a = 0, b = 0, c = 0, d = 0; pane.GetSelection(ref a, ref b, ref c, ref d);
            return (string)pane.CodeModule.Parent.Name + ":" + a + ":" + b + ":" + c + ":" + d;
        }
        /// <summary>Actualise l’état d’exécution affiché dans Monaco lorsque change le mode du projet.</summary>
        /// <returns>Tâche terminée après la mise à jour de l’état d’exécution.</returns>
private async Task ObserveDebugMode()
        {
            if (busy || Current == null || !(Current.Module is EditorVbeModule native)) return;
            int mode = (int)((dynamic)native.Project).Mode;
            if (mode != 1)
            {
                if (lastDebugMode != mode) foreach (var doc in documents.Values.ToArray()) await Script("execution", doc.Id, 0, false);
                lastDebugMode = mode; lastDebugPosition = null; return;
            }
            string position = DebugPosition(native);
            if (lastDebugMode == mode && lastDebugPosition == position)
            {
                if (lastDebugDocument != null && documents.TryGetValue(lastDebugDocument, out var doc) &&
                    !doc.Dirty && !doc.Conflict && lastExecutionVersion != versions[doc.Id])
                { await Script("execution", doc.Id, lastExecutionLine, false); lastExecutionVersion = versions[doc.Id]; }
                return;
            }
            await EditorCommand(new EditorMessage { id = Current.Id, version = versions[Current.Id], name = "show_next_statement" });
            lastDebugMode = mode; lastDebugPosition = DebugPosition(native);
        }
        /// <summary>Valide puis exécute une commande de compilation ou de débogage sur le projet courant.</summary>
        /// <param name="message">Commande et révision envoyées par l’interface Monaco.</param>
        /// <returns>Tâche terminée après l’exécution et l’actualisation de l’interface.</returns>
        /// <exception cref="InvalidOperationException">Le document a changé, le brouillon n’est pas synchronisé ou la commande native est indisponible.</exception>
        /// <exception cref="ArgumentException">Le nom de commande n’est pas reconnu.</exception>
private async Task EditorCommand(EditorMessage message)
        {
            while (busy && !closing && !IsDisposed) await Task.Delay(15);
            if (closing || IsDisposed || !documents.TryGetValue(message.id ?? "", out var document) || !(document.Module is EditorVbeModule native)) return;
            busy = true;
            try
            {
                bool observation = message.name == "show_next_statement";
                await CaptureDocuments();
                if (!observation)
                {
                    if (versions[document.Id] != message.version) throw new InvalidOperationException("The editor changed before the command. Retry at the current location.");
                    await ProcessDocumentsCore(true);
                    if (document.Dirty || document.Conflict) throw new InvalidOperationException("Synchronize or resolve the draft before compiling or debugging.");
                }
                await Task.Yield(); // Native commands must never execute inside a WebView callback.
                var debugger = new VbeDebug(native.Vbe);
                int mode = (int)((dynamic)native.Project).Mode;
                if (message.name == "compile")
                {
                    if (mode != 2) throw new InvalidOperationException("Compilation requires design mode.");
                    native.ShowNative(1, 1);
                    int hostProcessId = native.HostProcessId;
                    EnsureCompileDialogAbsent(hostProcessId);
                    string diagnostic;
                    using (var completed = new ManualResetEventSlim())
                    {
                        var observe = Task.Run(() => ObserveCompileDialog(completed, hostProcessId));
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
                        if (doc.Dirty || doc.Conflict || EditorDocument.Normalize(target.Read()) != doc.Text)
                        { SetResultStatus(diagnostic); return; } // Never underline an uncompiled draft.
                        await Script("diagnostics", doc.Id, versions[doc.Id], new[] { new { message = diagnostic, startLineNumber = line, startColumn = column, endLineNumber = endLine, endColumn = Math.Max(column + 1, endColumn), severity = 8, source = "VBA compiler" } });
                        await Script("reveal", line, column);
                    }
                    else SetResultStatus(UiText.Get("Compilation finished: no native diagnostics observed. Macros were not executed."));
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
                    var doc = documents.Values.FirstOrDefault(item => item.Module is EditorVbeModule module && module.IsComponent(target.Component));
                    if (doc == null || doc != Current) doc = await OpenModule(target);
                    lastDebugDocument = doc.Id; lastExecutionLine = line; lastExecutionVersion = versions[doc.Id];
                    foreach (var item in documents.Values.ToArray()) await Script("execution", item.Id, item == doc && !doc.Dirty && !doc.Conflict ? line : 0);
                }
                else
                {
                    if (!new[] { "toggle_breakpoint", "step_into", "step_over", "step_out" }.Contains(message.name)) throw new ArgumentException("Unknown editor command.");
                    object control = null;
                    native.ShowNative(Math.Max(1, message.line), 1);
                    ((dynamic)native.Vbe).ActiveCodePane.Window.SetFocus();
                    // Search the native command inventory, never emit global keyboard shortcuts.
                    for (int offset = 0; offset < 2000 && control == null; offset += 200)
                    {
                        var page = ((IEnumerable)debugger.ListCommands(null, offset, 200)).Cast<object>().ToArray();
                        foreach (dynamic item in page)
                            if ((bool)item.Enabled && (message.name != "toggle_breakpoint" || (int)item.Id == 51) && VbeDebug.IsAllowed(message.name, (string)item.Caption, mode)) { control = item; break; }
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
                if (!observation) { BringToFront(); Browser?.Focus(); }
            }
            finally { busy = false; }
        }
    }
}
