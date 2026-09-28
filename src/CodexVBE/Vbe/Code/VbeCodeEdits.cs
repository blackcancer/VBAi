using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;

namespace CodexVBE
{
    internal sealed class VbeCodeEdits
    {
        private readonly Func<Request, Response> execute;
        private readonly List<Entry> undo = new List<Entry>();
        private readonly List<Entry> redo = new List<Entry>();
        private bool replaying;
        private sealed class Entry { internal string Project, Module, Before, After; }
        internal VbeCodeEdits(Func<Request, Response> execute) { this.execute = execute; }
        internal void Record(string project, string module, string before, string after)
        {
            if (replaying || before == after) return;
            undo.Add(new Entry { Project = project, Module = module, Before = before, After = after });
            redo.Clear();
            // Bounded, session-local history. Never pretend this is the native VBE undo stack.
            while (undo.Count > 50 || undo.Sum(x => (long)x.Before.Length + x.After.Length) > 4 * 1024 * 1024) undo.RemoveAt(0);
        }
        internal object Edit(Request request, bool preview)
        {
            string before = Read(request.Project, request.Module);
            Check(before, request.ExpectedSha256);
            string after = VbaTextEdits.Transform(before, request);
            if (preview) return new { request.Project, request.Module, Before = before, After = after,
                ExpectedSha256 = Hash(before), Changed = before != after,
                Scope = "Explicit module line range; identifier replacements are lexical, not semantic refactoring." };
            return Write(request.Project, request.Module, before, after);
        }
        internal object Replay(Request request, bool forward)
        {
            var source = forward ? redo : undo; var destination = forward ? undo : redo;
            var entry = source.LastOrDefault(x => string.Equals(x.Project, request.Project, StringComparison.OrdinalIgnoreCase) && string.Equals(x.Module, request.Module, StringComparison.OrdinalIgnoreCase));
            if (entry == null) throw new InvalidOperationException("No VBAi code edit is available for this module in this session.");
            string current = Read(request.Project, request.Module);
            Check(current, request.ExpectedSha256);
            Check(current, Hash(forward ? entry.Before : entry.After));
            replaying = true;
            try {
                object result = Write(request.Project, request.Module, current, forward ? entry.After : entry.Before);
                source.Remove(entry); destination.Add(entry); return result;
            } finally { replaying = false; }
        }
        private object Write(string project, string module, string before, string after)
        {
            Response result = execute(new Request { Command = "replace_lines", Project = project, Module = module,
                ExpectedSha256 = Hash(before), StartLine = 1, Count = CodeRollback.Lines(before).Length, Text = after });
            if (!result.Ok) throw new InvalidOperationException(result.Error);
            string readback = Read(project, module);
            if (!string.Equals(readback, after, StringComparison.Ordinal))
                throw new InvalidOperationException("VBE text differs from the requested edit. Read the module before continuing.");
            return result.Data;
        }
        private string Read(string project, string module)
        {
            Response result = execute(new Request { Command = "read_module", Project = project, Module = module });
            if (!result.Ok) throw new InvalidOperationException(result.Error);
            return (string)((dynamic)result.Data).Code;
        }
        private static void Check(string code, string expected)
        {
            if (string.IsNullOrWhiteSpace(expected) || !string.Equals(Hash(code), expected, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The module changed. Read its current code and SHA before editing or replaying history.");
        }
        private static string Hash(string text)
        { using (var sha = SHA256.Create()) return BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(text))).Replace("-", "").ToLowerInvariant(); }
    }
}
