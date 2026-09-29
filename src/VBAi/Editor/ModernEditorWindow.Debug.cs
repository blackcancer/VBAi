using System;
using System.Collections;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace VBAi
{
    /// <summary>Suit le mode d’exécution VBE et exécute les commandes de compilation ou de débogage native.</summary>
    internal sealed partial class ModernEditorWindow
    {
        /// <summary>Dernier mode VBE observé pour le document actif.</summary>
        private int lastDebugMode = -1;
        private readonly SemaphoreSlim debugCommands = new SemaphoreSlim(1, 1);
        private bool observingDebug;
        /// <summary>Orders user/native commands without collapsing repeated toggles or steps.</summary>
        private async Task EditorCommand(EditorMessage message)
        {
            var timing = PerformanceSample == null ? null : System.Diagnostics.Stopwatch.StartNew();
            await debugCommands.WaitAsync();
            try { await ExecuteEditorCommand(message); }
            finally
            {
                debugCommands.Release(); Measure("debug." + message.name, timing);
                if (message.request > 0 && Ready && !closing && !IsDisposed) await Script("commandFinished", message.request);
            }
        }
        /// <summary>Updates all execution decorations in one renderer round trip.</summary>
        private Task ExecutionState(string id, int line, bool reveal = true)
        { return Script("executionBatch", id, line, reveal); }
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
            if (VbeDebugInspection.IsActive || busy || Current == null || !(Current.Module is EditorVbeModule native)) return;
            int mode = (int)((dynamic)native.Project).Mode;
            if (mode != 1)
            {
                if (lastDebugMode != mode) await ExecutionState(null, 0, false);
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
        private async Task ExecuteEditorCommand(EditorMessage message)
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
                    if (message.name == "toggle_breakpoint" && !VbaBreakpointLocation.CanRequest(document.Text, message.line)) return;
                    await ProcessCapturedDocumentsCore(true, false, onlyDocument: message.name == "toggle_breakpoint" ? document : null);
                    if ((message.name != "compile" && versions[document.Id] != message.version) || document.Dirty || document.Conflict) throw new InvalidOperationException("Synchronize or resolve the draft before compiling or debugging.");
                }
                await Task.Yield(); // Native commands must never execute inside a WebView callback.
                if (observation && VbeDebugInspection.IsActive) return;
                if (closing || IsDisposed || !documents.ContainsKey(document.Id)) return;
                if (!observation && (document.Dirty || document.Conflict || (message.name != "compile" && versions[document.Id] != message.version)))
                    throw new InvalidOperationException("The editor changed before the command. Retry at the current location.");
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
                    await ExecutionState(doc.Id, !doc.Dirty && !doc.Conflict ? line : 0);
                }
                else
                {
                    if (!new[] { "toggle_breakpoint", "step_into", "step_over", "step_out" }.Contains(message.name)) throw new ArgumentException("Unknown editor command.");
                    string source = native.Read();
                    if (message.name == "toggle_breakpoint" && !VbaBreakpointLocation.CanRequest(source, message.line)) return;
                    native.ShowNative(Math.Max(1, message.line), 1);
                    ((dynamic)native.Vbe).ActiveCodePane.Window.SetFocus();
                    object control = debugger.FindEditorCommand(message.name, mode);
                    if (control == null) throw new InvalidOperationException("The native debug command is unavailable in the current mode.");
                    debugger.InvokeCommand(new Request { Project = native.ProjectName, Module = native.ModuleName, ExpectedMode = mode,
                        ExpectedSha256 = EditorDocument.Hash(source), StartLine = Math.Max(1, message.line), Action = message.name, ControlId = (int)((dynamic)control).Id, ControlCaption = (string)((dynamic)control).Caption });
                    // VBIDE cannot enumerate breakpoints. This is deliberately a hollow request marker.
                    if (message.name == "toggle_breakpoint") await Script("breakpointRequested", document.Id, message.line);
                    else
                    {
                        lastDebugMode = -1; await ExecutionState(null, 0);
                        if (Ready && IsHandleCreated) BeginInvoke(new Action(() => DebugTimerTick(this, EventArgs.Empty)));
                    }
                }
                if (!observation) { BringToFront(); Browser?.Focus(); }
            }
            finally { busy = false; }
        }
    }
}
