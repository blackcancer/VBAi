using System;
using System.Linq;
using System.Threading.Tasks;

namespace VBAi
{

    /// <summary>Répond aux demandes de symboles et d’ouverture de définition de l’éditeur Monaco.</summary>
    internal sealed partial class ModernEditorWindow
    {

        /// <summary>Language work never queues behind draft persistence and synchronization.</summary>
        private EditorSyncWorker languageWorker;

        /// <summary>Worker-owned bounded cache of parsed module and referenced-library declarations.</summary>
        private readonly EditorLanguageCache languageCache = new EditorLanguageCache();

        /// <summary>Latest cancellable completion or hover request per renderer document ID.</summary>
        private readonly System.Collections.Generic.Dictionary<string, System.Threading.CancellationTokenSource> languageRequests = new System.Collections.Generic.Dictionary<string, System.Threading.CancellationTokenSource>();

        /// <summary>Serializes language requests so newer requests can cancel older queued work.</summary>
        private readonly System.Threading.SemaphoreSlim languageGate = new System.Threading.SemaphoreSlim(1, 1);

        /// <summary>Project identity whose native source snapshot is currently cached.</summary>
        private object languageSourceProject;

        /// <summary>Last captured native source set, reused for less than one second between requests.</summary>
        private EditorSource[] languageSources;

        /// <summary>Non-broken referenced library paths captured with <see cref="languageSources"/>.</summary>
        private string[] languageReferencePaths;

        /// <summary>Monotonic timestamp in milliseconds when the cached native source set was captured.</summary>
        private long languageSourcesAt;

        /// <summary>Replaceable monotonic millisecond clock used to bound native catalog snapshot reuse.</summary>
        internal Func<long> LanguageClock = () => System.Diagnostics.Stopwatch.GetTimestamp() / (System.Diagnostics.Stopwatch.Frequency / 1000);

        /// <summary>Bounds native catalog reads during a burst of completion/hover requests.</summary>
        /// <param name="native">Native module providing the project whose source and references are indexed.</param>
        /// <param name="cancellation">Cancellation requested when the renderer request is superseded or native commands become unavailable.</param>
        /// <returns>Captured module sources, reusing the same project's snapshot for up to one second.</returns>
        private Task<EditorSource[]> ReadLanguageSources(EditorVbeModule native, System.Threading.CancellationToken cancellation)
        {
            return VbeUiTask.Run(async () =>
            {
                long now = LanguageClock();
                if (!ReferenceEquals(languageSourceProject, native.Project) || languageSources == null || now - languageSourcesAt >= 1000)
                {
                    var captured = await native.Sources(cancellation, () => debugCommands.CurrentCount == 0);
                    var paths = ReadLanguageReferencePaths(native.Project, cancellation,
                        () => debugCommands.CurrentCount == 0);
                    languageSourceProject = native.Project; languageSources = captured;
                    languageReferencePaths = paths; languageSourcesAt = LanguageClock();
                }
                // Overlay drafts on a copy; never poison the cached native snapshot.
                return languageSources.Select(source => new EditorSource { Module = source.Module, Text = source.Text, ComponentType = source.ComponentType }).ToArray();
            });
        }

        /// <summary>Captures scalar library paths on the project owner without retaining native enumeration objects.</summary>
        /// <param name="project">Borrowed project held by the native module lease.</param>
        /// <param name="cancellation">Cancellation checked before native acquisition and for each item.</param>
        /// <param name="nativeUnavailable">Owner-thread check that rejects reads while native commands are occupied.</param>
        /// <param name="releaseAcquired">Optional acquisition-release boundary for ownership regressions.</param>
        /// <param name="indexedCollection">Optional one-based collection classification boundary.</param>
        /// <returns>Detached paths of non-broken references, preserving the native order.</returns>
        internal static string[] ReadLanguageReferencePaths(object project,
            System.Threading.CancellationToken cancellation, Func<bool> nativeUnavailable,
            Action<object> releaseAcquired = null, Func<object, bool> indexedCollection = null)
        {
            if (project == null) throw new ArgumentNullException(nameof(project));
            if (nativeUnavailable == null) throw new ArgumentNullException(nameof(nativeUnavailable));
            void Guard()
            {
                cancellation.ThrowIfCancellationRequested();
                if (nativeUnavailable()) throw new OperationCanceledException();
            }
            Guard();
            var paths = new System.Collections.Generic.List<string>();
            using (var references = indexedCollection == null
                ? new VbaTestComReferences(releaseAcquired)
                : new VbaTestComReferences(releaseAcquired, indexedCollection))
            foreach (dynamic reference in references.Items((object)((dynamic)project).References))
            {
                Guard();
                bool broken = (bool)reference.IsBroken;
                Guard();
                if (!broken)
                {
                    string path = (string)reference.FullPath;
                    // Native getters can pump the owner thread and change authority.
                    Guard();
                    paths.Add(path);
                }
            }
            Guard();
            return paths.ToArray();
        }

        /// <summary>Reads fresh project state but only parses and transfers changed language data.</summary>
        /// <param name="message">Completion/hover request with document revision, request key, and optional compact-stream state.</param>
        /// <returns>A task completed after a response is sent when the window remains ready, or after cancellation/staleness is discarded.</returns>
        private Task LanguageRequest(EditorMessage message)
        {
            return VbeUiTask.RunAsync(async () =>
            {
                object response = null;
                var cancellation = new System.Threading.CancellationTokenSource();
                string id = message.id ?? "";
                bool entered = false;
                if (languageRequests.TryGetValue(id, out var previous)) previous.Cancel();
                languageRequests[id] = cancellation;
                try
                {
                    await languageGate.WaitAsync(cancellation.Token); entered = true;
                    if (debugCommands.CurrentCount == 0) return;
                    await CaptureDocuments();
                    cancellation.Token.ThrowIfCancellationRequested();
                    if (!documents.TryGetValue(id, out var doc) || versions[doc.Id] != message.version || closing || IsDisposed) return;
                    var native = doc.Module as EditorVbeModule;
                    var sources = native == null ? documents.Values.Where(d => !(d.Module is EditorVbeModule)).Select(d => new EditorSource { Module = d.Module.Name, Text = d.Text, ComponentType = 1 }).ToArray() : await ReadLanguageSources(native, cancellation.Token);
                    cancellation.Token.ThrowIfCancellationRequested();
                    foreach (var entry in documents.Values)
                    {
                        var other = entry.Module as EditorVbeModule;
                        if (native != null && (other == null || !object.ReferenceEquals(native.Project, other.Project))) continue;
                        var source = sources.FirstOrDefault(s => s.Module == (other?.ModuleName ?? entry.Module.Name));
                        if (source != null) source.Text = entry.Text;
                    }
                    string[] libraryPaths = native == null ? Array.Empty<string>() : languageReferencePaths;
                    if (languageWorker == null) languageWorker = new EditorSyncWorker("VBAi editor language");
                    var snapshot = await languageWorker.Evaluate(() => languageCache.Build(sources, libraryPaths, cancellation.Token));
                    cancellation.Token.ThrowIfCancellationRequested();
                    if (!documents.ContainsKey(id) || versions[id] != message.version || closing || IsDisposed) return;
                    string key = doc.Id + ":" + snapshot.Key + (message.includeSources ? ":sources" : "");
                    if (message.compact && message.knownLanguage == key) response = new { unchanged = true, key };
                    else if (message.compact) response = new
                    {
                        id = doc.Id,
                        module = native?.ModuleName ?? doc.Module.Name,
                        key,
                        parts = snapshot.Parts.Select(part => new
                        {
                            name = part.Name,
                            key = part.Key,
                            symbols = message.knownParts != null && message.knownParts.TryGetValue(part.Name, out var known) && known == part.Key ? null : part.Symbols
                        }).ToArray(),
                        sources = message.includeSources ? sources : Array.Empty<EditorSource>()
                    };
                    else response = new { id = doc.Id, module = native?.ModuleName ?? doc.Module.Name, key, symbols = snapshot.Symbols, sources };
                }
                catch (OperationCanceledException) { }
                catch (Exception error) { LoadLog.Write("Monaco language service: " + error.Message); }
                finally
                {
                    if (entered) languageGate.Release();
                    if (languageRequests.TryGetValue(id, out var current) && current == cancellation) languageRequests.Remove(id);
                    cancellation.Dispose();
                    if (Ready && !closing && !IsDisposed) await Script("languageReply", message.request, response);
                }
            });
        }

        /// <summary>Ouvre le module qui définit le symbole demandé puis révèle sa ligne et sa colonne.</summary>
        /// <param name="message">Destination du symbole fournie par l’index de langage.</param>
        /// <returns>Tâche terminée après l’ouverture et le positionnement éventuels.</returns>
        private Task OpenDefinition(EditorMessage message)
        {
            return VbeUiTask.RunAsync(async () =>
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
            });
        }
    }
}
