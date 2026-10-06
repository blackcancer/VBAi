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

        /// <summary>Maintains the modules state for editor language cache.</summary>
        private readonly Dictionary<string, Entry> modules = new Dictionary<string, Entry>();

        /// <summary>Maintains the retained characters state for editor language cache.</summary>
        private int retainedCharacters;

        /// <summary>Number of modules actually parsed, for performance diagnostics.</summary>
        /// <value>Current parsed modules exposed by editor language cache.</value>
        internal int ParsedModules { get; private set; }

        /// <summary>Owns the entry state and operations.</summary>
        private sealed class Entry
        {

            /// <summary>Maintains the symbols state for entry.</summary>
            internal EditorSymbol[] Symbols;

            /// <summary>Maintains the receivers state for entry.</summary>
            internal string[] Receivers;
        }

        /// <summary>A module or reference symbol bucket independently versioned for transport.</summary>
        internal sealed class Part
        {

            /// <summary>Maintains the name and key state for part.</summary>
            internal string Name, Key;

            /// <summary>Maintains the symbols state for part.</summary>
            internal EditorSymbol[] Symbols;
        }

        /// <summary>Immutable result identified by all source and reference revisions.</summary>
        internal sealed class Snapshot
        {

            /// <summary>Maintains the key state for snapshot.</summary>
            internal string Key;

            /// <summary>Maintains the symbols state for snapshot.</summary>
            internal EditorSymbol[] Symbols;

            /// <summary>Maintains the parts state for snapshot.</summary>
            internal Part[] Parts;
        }

        /// <summary>Builds only changed module declarations and reuses reference metadata on its owning worker.</summary>
        /// <param name="sources">editor source[] that supplies the sources for this operation.</param>
        /// <param name="paths">string[] that supplies the paths for this operation.</param>
        /// <param name="cancellation">Token used to cancel the operation.</param>
        /// <returns>snapshot produced by the operation for build on editor language cache.</returns>
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
