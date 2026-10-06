using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;

namespace VBAi
{

    /// <summary>Bounded worker-owned declaration cache; every native snapshot is still validated by its caller.</summary>
    internal sealed class EditorLanguageCache
    {

        /// <summary>Worker-owned cache from a content-and-component hash to parsed declarations and receiver names.</summary>
        private readonly Dictionary<string, Entry> modules = new Dictionary<string, Entry>();

        /// <summary>Total source characters retained by <see cref="modules"/> for its 8 MiB cache bound.</summary>
        private int retainedCharacters;

        /// <summary>Number of modules actually parsed, for performance diagnostics.</summary>
        /// <value>Current parsed modules exposed by editor language cache.</value>
        internal int ParsedModules { get; private set; }

        /// <summary>Parsed symbols and receiver expressions cached for one module-content key.</summary>
        private sealed class Entry
        {

            /// <summary>Declarations produced by the language index for this module revision.</summary>
            internal EditorSymbol[] Symbols;

            /// <summary>Distinct identifier receivers preceding member-access dots in this module revision.</summary>
            internal string[] Receivers;
        }

        /// <summary>A module or reference symbol bucket independently versioned for transport.</summary>
        internal sealed class Part
        {

            /// <summary>Transport label and content key identifying this module or reference bucket.</summary>
            internal string Name, Key;

            /// <summary>Declarations supplied by this independently versioned bucket.</summary>
            internal EditorSymbol[] Symbols;
        }

        /// <summary>Immutable result identified by all source and reference revisions.</summary>
        internal sealed class Snapshot
        {

            /// <summary>Hash of the ordered module and reference revisions represented by this snapshot.</summary>
            internal string Key;

            /// <summary>Combined declarations from captured modules and selected referenced libraries.</summary>
            internal EditorSymbol[] Symbols;

            /// <summary>Module and reference buckets that let the editor stream only changed portions.</summary>
            internal Part[] Parts;
        }

        /// <summary>Builds only changed module declarations and reuses reference metadata on its owning worker.</summary>
        /// <param name="sources">Ordered module snapshots to parse or reuse from the worker-local cache.</param>
        /// <param name="paths">Type-library paths whose metadata may provide external symbols.</param>
        /// <param name="cancellation">Token used to cancel the operation.</param>
        /// <returns>A snapshot keyed by source content and reference metadata, containing combined symbols and per-source parts.</returns>
        internal Snapshot Build(EditorSource[] sources, string[] paths, CancellationToken cancellation)
        {
            var symbols = new List<EditorSymbol>();
            var parts = new List<Part>();
            var receivers = new List<string>();
            var signature = new StringBuilder();
            foreach (var source in sources)
            {
                cancellation.ThrowIfCancellationRequested();
                string key = EditorDocument.Hash(source.Module.Length + ":" + source.Module + ":" + source.ComponentType + ":" + source.Text);
                signature.Append(key);
                if (!modules.TryGetValue(key, out var entry))
                {
                    entry = new Entry { Symbols = EditorLanguageIndex.Build(new[] { source }),
                        Receivers = Regex.Matches(source.Text, @"\b([\p{L}_][\p{L}\p{N}_]*)\s*\.").Cast<Match>().Select(m => m.Groups[1].Value).Distinct(StringComparer.OrdinalIgnoreCase).ToArray() };
                    ParsedModules++;
                    if (modules.Count >= 256 || retainedCharacters + source.Text.Length > 8 * 1024 * 1024) { modules.Clear(); retainedCharacters = 0; }
                    if (source.Text.Length <= 8 * 1024 * 1024) { modules[key] = entry; retainedCharacters += source.Text.Length; }
                }
                symbols.AddRange(entry.Symbols); receivers.AddRange(entry.Receivers);
                parts.Add(new Part { Name = "m:" + source.Module, Key = key, Symbols = entry.Symbols });
            }
            var referenceSignature = new StringBuilder();
            foreach (string path in paths)
                referenceSignature.Append(path.Length).Append(':').Append(path).Append(':').Append(System.IO.File.GetLastWriteTimeUtc(path).Ticks).Append(';');
            cancellation.ThrowIfCancellationRequested();
            if (paths.Length > 0)
            {
                var types = symbols.Where(s => !string.IsNullOrEmpty(s.TypeName)).Select(s => s.TypeName).Concat(receivers)
                    .Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(t => t, StringComparer.OrdinalIgnoreCase).ToArray();
                var external = EditorReferenceIndex.Read(paths, types);
                symbols.AddRange(external);
                foreach (var type in types) referenceSignature.Append(type.Length).Append(':').Append(type);
                parts.Add(new Part { Name = "r:", Key = EditorDocument.Hash(referenceSignature.ToString()), Symbols = external });
            }
            signature.Append(referenceSignature);
            cancellation.ThrowIfCancellationRequested();
            return new Snapshot { Key = EditorDocument.Hash(signature.ToString()), Symbols = symbols.ToArray(), Parts = parts.ToArray() };
        }
    }
}
