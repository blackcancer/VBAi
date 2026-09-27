using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace CodexVBE
{
    // All methods run on the VBE UI thread. Never send COM objects to the Git worker.
    internal sealed class VbaGitProject
    {
        private readonly Func<object> resolve;
        private readonly string hostPath;
        [System.Runtime.InteropServices.DllImport("kernel32.dll")]
        private static extern uint GetACP();
        private static Encoding NativeEncoding { get { return Encoding.GetEncoding(
            (int)GetACP(), EncoderFallback.ExceptionFallback, DecoderFallback.ExceptionFallback); } }

        internal VbaGitProject(Func<object> resolve, string hostPath) { this.resolve = resolve; this.hostPath = hostPath; }

        internal void OpenModule(string name, int line = 1)
        {
            VbaGitSnapshot.ValidateName(name);
            dynamic project = CheckedProject();
            dynamic pane = project.VBComponents.Item(name).CodeModule.CodePane;
            pane.Show();
            pane.SetSelection(Math.Max(1, line), 1, Math.Max(1, line), 1);
        }

        private object CheckedProject()
        {
            dynamic project = resolve();
            if (!string.Equals(Path.GetFullPath((string)project.FileName), hostPath, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException(UiText.Get("The linked document changed. Reopen GitHub integration."));
            if ((int)project.Mode != 2 || (int)project.Protection != 0)
                throw new InvalidOperationException(UiText.Get("The VBA project must be unlocked and in design mode."));
            return project;
        }

        internal VbaGitSnapshot Capture()
        {
            dynamic project = CheckedProject();
            var components = new List<VbaGitComponent>();
            var files = new Dictionary<string, byte[]>(StringComparer.Ordinal);
            using (var scratch = new Scratch())
            {
                foreach (dynamic item in project.VBComponents)
                {
                    var component = new VbaGitComponent { Name = (string)item.Name, Type = (int)item.Type };
                    // Validate before using a COM-supplied name as a path.
                    VbaGitSnapshot.ValidateName(component.Name);
                    string file = Path.Combine(scratch.Path, component.FileName);
                    if (component.Type == 100)
                        files.Add(component.FileName, VbaGitSnapshot.Utf8.GetBytes(Code(item.CodeModule)));
                    else
                    {
                        item.Export(file);
                        files.Add(component.FileName, VbaGitSnapshot.Utf8.GetBytes(Normalize(File.ReadAllText(file, NativeEncoding))));
                        string frx = Path.Combine(scratch.Path, component.Name + ".frx");
                        component.HasResources = component.Type == 3 && File.Exists(frx);
                        if (component.HasResources) files.Add(component.Name + ".frx", File.ReadAllBytes(frx));
                    }
                    components.Add(component);
                }
            }
            var references = new List<string>();
            foreach (dynamic reference in project.References)
            {
                if ((bool)reference.IsBroken) throw new InvalidOperationException(UiText.Get("Missing VBA reference: fix it before synchronizing."));
                references.Add(((string)reference.GUID).ToUpperInvariant() + ":" + (int)reference.Major + ":" + (int)reference.Minor);
            }
            return new VbaGitSnapshot(new VbaGitManifest {
                Components = components.OrderBy(x => x.Name, StringComparer.Ordinal).ToArray(),
                References = string.Join(";", references.OrderBy(x => x, StringComparer.Ordinal))
            }, files);
        }

        internal void Apply(VbaGitSnapshot target, VbaGitSnapshot expected, Action beforeMutation = null)
        {
            if (!Capture().SameAs(expected)) throw new InvalidOperationException(UiText.Get("VBA changed during synchronization. No import performed."));
            if (target.Manifest.References != expected.Manifest.References)
                throw new InvalidOperationException(UiText.Get("VBA references differ. Align them in Tools > References before importing."));
            var beforeDocs = expected.Manifest.Components.Where(x => x.Type == 100).Select(x => x.Name);
            var afterDocs = target.Manifest.Components.Where(x => x.Type == 100).Select(x => x.Name);
            if (!beforeDocs.SequenceEqual(afterDocs))
                throw new InvalidOperationException(UiText.Get("Document modules do not match. Sheets and host modules must already exist with the same names."));

            using (var scratch = new Scratch())
            {
                // Encode and materialize the entire import before touching the live project.
                foreach (var file in target.Files)
                    File.WriteAllBytes(Path.Combine(scratch.Path, file.Key), file.Key.EndsWith(".frx", StringComparison.Ordinal) ? file.Value :
                        NativeEncoding.GetBytes(VbaGitSnapshot.Utf8.GetString(file.Value).Replace("\n", "\r\n")));
                dynamic project = CheckedProject();
                beforeMutation?.Invoke();
                var changed = new HashSet<string>(StringComparer.Ordinal);
                foreach (var old in expected.Manifest.Components)
                {
                    var next = target.Manifest.Components.FirstOrDefault(x => x.Name == old.Name);
                    bool same = next != null && next.Type == old.Type && next.HasResources == old.HasResources &&
                        expected.Files[old.FileName].SequenceEqual(target.Files[next.FileName]) &&
                        (!old.HasResources || expected.Files[old.Name + ".frx"].SequenceEqual(target.Files[next.Name + ".frx"]));
                    if (same) continue;
                    changed.Add(old.Name);
                    if (old.Type != 100) project.VBComponents.Remove(project.VBComponents.Item(old.Name));
                }
                foreach (var next in target.Manifest.Components)
                {
                    if (!changed.Contains(next.Name) && expected.Manifest.Components.Any(x => x.Name == next.Name)) continue;
                    if (next.Type == 100)
                    {
                        dynamic module = project.VBComponents.Item(next.Name).CodeModule;
                        int count = (int)module.CountOfLines;
                        if (count > 0) module.DeleteLines(1, count);
                        string code = VbaGitSnapshot.Utf8.GetString(target.Files[next.FileName]);
                        if (code.Length > 0) module.InsertLines(1, code.Replace("\n", "\r\n"));
                    }
                    else
                    {
                        // A failing COM call can still have applied. Do not retry or automatically re-import.
                        dynamic imported = project.VBComponents.Import(Path.Combine(scratch.Path, next.FileName));
                        if ((string)imported.Name != next.Name || (int)imported.Type != next.Type)
                            throw new InvalidOperationException(UiText.Get("Unexpected identity after import: ") + next.Name + UiText.Get(". Use Restore."));
                    }
                }
            }
            if (!Capture().SameAs(target)) throw new InvalidOperationException(UiText.Get("The VBE did not preserve the imported sources exactly. Use Restore or check the project."));
        }

        private static string Code(dynamic module)
        {
            int count = (int)module.CountOfLines;
            return count == 0 ? "" : Normalize((string)module.Lines[1, count]);
        }
        private static string Normalize(string text) { return text.Replace("\r\n", "\n").Replace("\r", "\n"); }

        private sealed class Scratch : IDisposable
        {
            internal readonly string Path = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "CodexVBE", "GitTemporary", Guid.NewGuid().ToString("N"));
            internal Scratch() { Directory.CreateDirectory(Path); }
            public void Dispose()
            {
                // Only our freshly generated, private flat directory is cleaned up.
                try { foreach (string file in Directory.GetFiles(Path)) File.Delete(file); Directory.Delete(Path); }
                catch (IOException) { } catch (UnauthorizedAccessException) { }
            }
        }
    }
}
