using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;

namespace CodexVBE
{
    internal sealed partial class VbeCodeEdits
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
        /// <summary>Prévisualise ou applique un renommage local lié à une déclaration et à la plage VBIDE.</summary>
        internal object RenameLocal(Request request, bool preview)
        {
            string before = Read(request.Project, request.Module); Check(before, request.ExpectedSha256);
            Response catalog = execute(new Request { Command = "list_procedures", Project = request.Project, Module = request.Module });
            if (!catalog.Ok) throw new InvalidOperationException(catalog.Error);
            dynamic data = catalog.Data;
            if (!string.Equals((string)data.Sha256, request.ExpectedSha256, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The module changed before resolving the procedure.");
            dynamic selected = null;
            foreach (dynamic procedure in data.Procedures)
                if ((int)procedure.Kind == request.ProcKind && string.Equals((string)procedure.Name, request.Procedure, StringComparison.OrdinalIgnoreCase))
                { if (selected != null) throw new InvalidOperationException("The procedure is ambiguous."); selected = procedure; }
            if (selected == null) throw new InvalidOperationException("The exact procedure is absent.");
            string after = VbaLocalRename.Transform(before, request, (int)selected.BodyLine, (int)selected.EndLine);
            if (preview) return new { request.Project, request.Module, request.Procedure, Before = before, After = after,
                ExpectedSha256 = Hash(before), Changed = before != after,
                Scope = "One explicit local variable/constant in its VBIDE procedure range; members/types/labels/named arguments are excluded. Parameters, conditional code and project-wide refactoring are refused." };
            if (request.ExpectedMode != 2) throw new ArgumentException("ExpectedMode=2 is required to rename a local declaration.");
            Response state = execute(new Request { Command = "debug_state", Project = request.Project });
            if (!state.Ok || (int)((dynamic)state.Data).Mode != 2) throw new InvalidOperationException("Renaming requires design mode.");
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
