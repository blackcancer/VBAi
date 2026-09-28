using System;
using System.Linq;
using System.Threading.Tasks;

namespace CodexVBE
{
    internal sealed partial class ModernEditorWindow
    {
        private async Task LanguageRequest(EditorMessage message)
        {
            object response = null;
            try
            {
                await CaptureDocuments();
                if (!documents.TryGetValue(message.id ?? "", out var doc) || versions[doc.Id] != message.version) return;
                var native = doc.Module as EditorVbeModule;
                var sources = native == null ? documents.Values.Where(d => !(d.Module is EditorVbeModule)).Select(d => new EditorSource { Module = d.Module.Name, Text = d.Text, ComponentType = 1 }).ToArray() : await native.Sources();
                foreach (var entry in documents.Values)
                {
                    var other = entry.Module as EditorVbeModule;
                    if (native != null && (other == null || !object.ReferenceEquals(native.Project, other.Project))) continue;
                    var source = sources.FirstOrDefault(s => s.Module == (other?.ModuleName ?? entry.Module.Name));
                    if (source != null) source.Text = entry.Text;
                }
                if (synchronizationWorker == null) synchronizationWorker = new EditorSyncWorker();
                var symbols = await synchronizationWorker.Evaluate(() => EditorLanguageIndex.Build(sources));
                if (native != null)
                {
                    var paths = new System.Collections.Generic.List<string>();
                    foreach (dynamic reference in ((dynamic)native.Project).References)
                        if (!(bool)reference.IsBroken) paths.Add((string)reference.FullPath);
                    string[] libraryPaths = paths.ToArray();
                    string[] types = symbols.Where(s => !string.IsNullOrEmpty(s.TypeName)).Select(s => s.TypeName).Distinct().ToArray();
                    var external = await synchronizationWorker.Evaluate(() => EditorReferenceIndex.Read(libraryPaths, types));
                    symbols = symbols.Concat(external).ToArray();
                }
                response = new { id = doc.Id, module = native?.ModuleName ?? doc.Module.Name, symbols, sources };
            }
            catch (Exception error) { LoadLog.Write("Monaco language service: " + error.Message); }
            finally { if (Ready && !IsDisposed) await Script("languageReply", message.request, response); }
        }
        private async Task OpenDefinition(EditorMessage message)
        {
            if (!documents.TryGetValue(message.id ?? "", out var doc)) return;
            if (doc.Module is EditorVbeModule native) await OpenModule(native.Sibling(message.module));
            else
            {
                var target = documents.Values.FirstOrDefault(d => d.Module.Name == message.module && !(d.Module is EditorVbeModule));
                if (target == null) return;
                await OpenModule(target.Module);
            }
            await Script("reveal", Math.Max(1, message.line), Math.Max(1, message.column));
        }
    }
}
