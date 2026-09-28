using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;

namespace CodexVBE
{
    /// <summary>All methods are called on the owning VBE UI thread. Bind to COM identity, not a mutable module name.</summary>
    internal sealed class EditorVbeModule : IEditorModule
    {
        private readonly object vbe, project, component;
        private readonly string sessionKey = Guid.NewGuid().ToString("N");
        private object nativeWindow;
        internal EditorVbeModule(object vbe, object project, object component)
        { this.vbe = vbe; this.project = project; this.component = component; }
        public string Name => (string)((dynamic)project).Name + " · " + (string)((dynamic)component).Name;
        public string Key
        {
            get
            {
                string path = null;
                try { path = (string)((dynamic)project).FileName; } catch { }
                return (!string.IsNullOrWhiteSpace(path) && Path.IsPathRooted(path) ? Path.GetFullPath(path).ToUpperInvariant() : sessionKey) + "|" + (string)((dynamic)component).Name;
            }
        }
        internal bool IsComponent(object other) => Same(component, other);
        internal object Component => component;
        internal object Project => project;
        internal object Vbe => vbe;
        internal string ModuleName => (string)((dynamic)component).Name;
        internal string ProjectName
        {
            get { try { string path = ((dynamic)project).FileName; if (!string.IsNullOrEmpty(path)) return path; } catch { } return (string)((dynamic)project).Name; }
        }
        internal async System.Threading.Tasks.Task<EditorSource[]> Sources()
        {
            Validate(); var sources = new System.Collections.Generic.List<EditorSource>();
            foreach (dynamic item in ((dynamic)project).VBComponents)
            {
                dynamic code = item.CodeModule; int count = code.CountOfLines;
                sources.Add(new EditorSource { Module = (string)item.Name, ComponentType = (int)item.Type, Text = count == 0 ? "" : (string)code.Lines[1, count] });
                await System.Threading.Tasks.Task.Yield(); // Let the host process input between COM module reads.
            }
            return sources.ToArray();
        }
        internal EditorVbeModule Sibling(string name)
        {
            Validate(); foreach (object item in ((dynamic)project).VBComponents)
                if (string.Equals((string)((dynamic)item).Name, name, StringComparison.OrdinalIgnoreCase)) return new EditorVbeModule(vbe, project, item);
            throw new InvalidOperationException("The VBA module was removed.");
        }
        internal void EnsureNativeWindow()
        {
            Validate(); nativeWindow = ((dynamic)component).CodeModule.CodePane.Window;
            ((dynamic)nativeWindow).Visible = true;
        }
        internal void CloseNativeWindow()
        {
            if (nativeWindow == null) return;
            try { ((dynamic)nativeWindow).Close(); }
            catch (COMException) { } // The project or its backing pane may already have closed.
            finally { nativeWindow = null; }
        }
        private static bool Same(object a, object b)
        {
            if (ReferenceEquals(a, b)) return true;
            if (!Marshal.IsComObject(a) || !Marshal.IsComObject(b)) return false;
            IntPtr x = Marshal.GetIUnknownForObject(a), y = Marshal.GetIUnknownForObject(b);
            try { return x == y; } finally { Marshal.Release(x); Marshal.Release(y); }
        }
        private void Validate()
        {
            bool found = false;
            foreach (object p in ((dynamic)vbe).VBProjects) if (Same(p, project)) { found = true; break; }
            if (!found) throw new InvalidOperationException("The VBA project is closed. Your draft is preserved.");
            found = false;
            foreach (object c in ((dynamic)project).VBComponents) if (Same(c, component)) { found = true; break; }
            if (!found) throw new InvalidOperationException("The VBA module was removed. Your draft is preserved.");
        }
        public bool CanWrite
        { get { Validate(); return (int)((dynamic)project).Mode == 2 && (int)((dynamic)project).Protection == 0; } }
        public string Read()
        {
            Validate(); dynamic module = ((dynamic)component).CodeModule;
            int count = (int)module.CountOfLines;
            return count == 0 ? "" : (string)module.Lines[1, count];
        }
        public string Write(string expected, string text)
        { return WritePrepared(expected, text, null); }
        internal string WritePrepared(string expected, string text, EditorSyncPlan plan)
        {
            Validate(); if (!CanWrite) throw new InvalidOperationException("VBA is running, paused or unavailable. Your draft is preserved.");
            string before = Read();
            if (EditorDocument.Normalize(before) != expected) throw new InvalidOperationException("The module changed in VBA. Resolve the conflict first.");
            // Preserve host encoding rather than letting unsupported characters silently become '?'.
            var encoding = System.Text.Encoding.GetEncoding(System.Text.Encoding.Default.CodePage,
                System.Text.EncoderFallback.ExceptionFallback, System.Text.DecoderFallback.ExceptionFallback);
            encoding.GetBytes(text);
            dynamic module = ((dynamic)component).CodeModule;
            if (plan != null && (plan.Before != expected || plan.After != text)) throw new InvalidOperationException("Stale synchronization plan.");
            var edit = plan?.Patch ?? EditorDocument.Difference(expected, text);
            if (TryRewriteAttributedDeclaration(expected, text, edit, out string rewritten)) return rewritten;
            GuardProcedureAttributes(expected, edit);
            try
            {
                // ReplaceLine retains the surrounding procedure for a single-line edit.
                if (edit.Item2 == 1 && edit.Item3.IndexOf('\n') < 0 && edit.Item3.Length > 0) module.ReplaceLine(edit.Item1, edit.Item3);
                else
                {
                    if (edit.Item2 > 0) module.DeleteLines(edit.Item1, edit.Item2);
                    if (edit.Item3.Length > 0) module.InsertLines(edit.Item1, edit.Item3);
                }
                return Read();
            }
            catch (Exception failure)
            {
                try
                {
                    string current = Read();
                    var restoration = EditorDocument.Difference(current, before);
                    GuardProcedureAttributes(current, restoration);
                    if (restoration.Item2 == 1 && restoration.Item3.IndexOf('\n') < 0 && restoration.Item3.Length > 0) module.ReplaceLine(restoration.Item1, restoration.Item3);
                    else
                    {
                        if (restoration.Item2 > 0) module.DeleteLines(restoration.Item1, restoration.Item2);
                        if (restoration.Item3.Length > 0) module.InsertLines(restoration.Item1, restoration.Item3);
                    }
                    if (Read() != before) throw new InvalidOperationException("Source restoration mismatch.");
                }
                catch (Exception rollback) { throw new InvalidOperationException("Code synchronization and restoration failed. Your draft is preserved.", new AggregateException(failure, rollback)); }
                throw new InvalidOperationException("Code synchronization failed; the original source was restored.", failure);
            }
        }
        private bool TryRewriteAttributedDeclaration(string before, string after, Tuple<int, int, string> patch, out string result)
        {
            result = null;
            // Document modules and designers have extra host-owned metadata; never rebuild them here.
            if ((int)((dynamic)component).Type != 1 || patch.Item2 != 1 || patch.Item3.Contains("\n")) return false;
            string directory = Path.Combine(Path.GetTempPath(), "VBAi-attributes-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(directory); bool preserveBackup = false;
            try
            {
                string backup = Path.Combine(directory, "original.bas"), changed = Path.Combine(directory, "changed.bas"), verify = Path.Combine(directory, "verify.bas");
                ((dynamic)component).Export(backup);
                string original = File.ReadAllText(backup, System.Text.Encoding.Default);
                string replacement = EditorAttributeRewrite.Prepare(original, before, patch);
                if (replacement == null) return false;
                File.WriteAllText(changed, replacement, System.Text.Encoding.Default);
                dynamic code = ((dynamic)component).CodeModule;
                Action<string> load = path => { int count = code.CountOfLines; if (count > 0) code.DeleteLines(1, count); code.AddFromFile(path); };
                if (!CanWrite || EditorDocument.Normalize(Read()) != before) throw new InvalidOperationException("The module changed before attribute restoration.");
                try
                {
                    load(changed);
                    result = Read(); ((dynamic)component).Export(verify);
                    if (EditorDocument.Normalize(result).TrimEnd('\n') != after.TrimEnd('\n') || EditorAttributeRewrite.Metadata(File.ReadAllText(verify, System.Text.Encoding.Default)) != EditorAttributeRewrite.Metadata(replacement))
                        throw new InvalidOperationException("Attributed source readback mismatch.");
                    return true;
                }
                catch (Exception failure)
                {
                    try
                    {
                        load(backup);
                        File.Delete(verify); ((dynamic)component).Export(verify);
                        if (EditorDocument.Normalize(Read()).TrimEnd('\n') != before.TrimEnd('\n') || EditorAttributeRewrite.Metadata(File.ReadAllText(verify, System.Text.Encoding.Default)) != EditorAttributeRewrite.Metadata(original))
                            throw new InvalidOperationException("Attribute restoration mismatch.");
                    }
                    catch (Exception rollback) { preserveBackup = true; throw new InvalidOperationException("Attributed edit and restoration failed; original export retained at " + backup, new AggregateException(failure, rollback)); }
                    throw new InvalidOperationException("Attributed edit failed; original source and attributes restored.", failure);
                }
            }
            finally { if (!preserveBackup) Directory.Delete(directory, true); }
        }
        public void ShowNative(int line, int column)
        {
            Validate(); dynamic pane = ((dynamic)component).CodeModule.CodePane;
            int count = (int)pane.CodeModule.CountOfLines;
            line = Math.Max(1, Math.Min(Math.Max(1, count), line)); column = Math.Max(1, column);
            pane.Show(); pane.SetSelection(line, column, line, column);
        }
        private void GuardProcedureAttributes(string before, Tuple<int, int, string> edit)
        {
            // CodeModule hides procedure metadata. Deleting/recreating a declaration can
            // destroy that metadata; unsupported declaration edits remain guarded.
            string directory = Path.Combine(Path.GetTempPath(), "VBAi-editor-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(directory);
            try
            {
                string path = Path.Combine(directory, "module.bas");
                ((dynamic)component).Export(path);
                string source = File.ReadAllText(path, System.Text.Encoding.Default);
                EnsureAttributeDeclarationsUntouched(source, before, edit);
            }
            finally { Directory.Delete(directory, true); }
        }
        internal static void EnsureAttributeDeclarationsUntouched(string exported, string before, Tuple<int, int, string> edit)
        {
            var names = new System.Collections.Generic.HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (System.Text.RegularExpressions.Match match in System.Text.RegularExpressions.Regex.Matches(exported, @"(?im)^\s*Attribute\s+([^\s.]+)\.")) names.Add(match.Groups[1].Value);
            if (names.Count == 0) return;
            foreach (var statement in VbaDeclarationIndex.Statements(before))
            {
                int p = 0;
                while (p < statement.Count && new[] { "public", "private", "friend", "static" }.Contains(statement[p].Text.ToLowerInvariant())) p++;
                if (p >= statement.Count) continue;
                string kind = statement[p].Text.ToLowerInvariant();
                int n = p + (kind == "property" ? 2 : 1);
                if ((kind != "sub" && kind != "function" && kind != "property") || n >= statement.Count || !names.Contains(statement[n].Text)) continue;
                int first = statement[0].Line, last = statement[statement.Count - 1].Line;
                // A zero-length insertion before the declaration does not destroy its metadata.
                bool intersects = edit.Item2 > 0 && edit.Item1 <= last && edit.Item1 + edit.Item2 - 1 >= first;
                bool splits = edit.Item2 == 0 && edit.Item1 > first && edit.Item1 <= last;
                if (intersects || splits) throw new InvalidOperationException("This change replaces a declaration with hidden procedure attributes. Edit its body or use an exported module to preserve its metadata.");
            }
        }
    }
}
