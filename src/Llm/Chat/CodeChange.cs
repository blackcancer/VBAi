using System;
using System.Linq;
using System.Text;

namespace CodexVBE
{
    internal enum CodeDiffKind { Context, Removed, Added, Notice }

    internal sealed class CodeDiffLine
    {
        public CodeDiffKind Kind { get; set; }
        public int? OldLine { get; set; }
        public int? NewLine { get; set; }
        public string Text { get; set; }
        public string Sign { get { return Kind == CodeDiffKind.Added ? "+" : Kind == CodeDiffKind.Removed ? "−" : ""; } }
    }

    // A session-local record of an actual VBIDE code edit. VBA remains the source of truth.
    internal sealed class CodeChange
    {
        public string Project { get; set; }
        public string Module { get; set; }
        public string Before { get; set; }
        public string After { get; set; }
        public string BeforeSha256 { get; set; }
        public string AfterSha256 { get; set; }
        public int AfterLineCount { get; set; }
        public DateTime Time { get; set; }
        public bool Restored { get; set; }
        public string TurnId { get; set; }
        public System.Collections.Generic.List<int> RestoredHunks { get; set; } = new System.Collections.Generic.List<int>();
        [System.Web.Script.Serialization.ScriptIgnore] public string Diff { get { return FormatDiff(Before, After); } }
        [System.Web.Script.Serialization.ScriptIgnore] public CodeDiffLine[] Rows { get { return BuildRows(Before, After); } }
        [System.Web.Script.Serialization.ScriptIgnore] public string Label { get { return Time.ToString("HH:mm:ss") + "  " + Project + "." + Module +
            (Restored ? "  (restauré)" : ""); } }
        public override string ToString() { return Label; }

        public CodeChange() { }
        public CodeChange(string project, string module, string before, string beforeSha256,
            string after, string afterSha256, int afterLineCount)
        {
            Project = project;
            Module = module;
            Before = before;
            BeforeSha256 = beforeSha256;
            After = after;
            AfterSha256 = afterSha256;
            AfterLineCount = afterLineCount;
            Time = DateTime.Now;
        }

        public static string Preview(string before, Request request)
        {
            return FormatDiff(before, ProposedCode(before, request));
        }

        public static CodeDiffLine[] PreviewRows(string before, Request request)
        {
            return BuildRows(before, ProposedCode(before, request));
        }

        private static string ProposedCode(string before, Request request)
        {
            string[] lines = Lines(before);
            if (request.StartLine < 1 || request.Count < 0 || request.StartLine > lines.Length + 1 ||
                request.Count > lines.Length - request.StartLine + 1 || request.Text == null)
                throw new ArgumentException("La plage de remplacement est invalide.");
            string[] inserted = Lines(request.Text);
            string[] result = lines.Take(request.StartLine - 1).Concat(inserted)
                .Concat(lines.Skip(request.StartLine - 1 + request.Count)).ToArray();
            return string.Join("\n", result);
        }

        public static string FormatDiff(string before, string after)
        {
            var rows = BuildRows(before, after);
            if (rows.Length == 1 && rows[0].Kind == CodeDiffKind.Notice) return rows[0].Text;
            var diff = new StringBuilder();
            foreach (var row in rows)
                diff.Append(row.Kind == CodeDiffKind.Added ? '+' :
                    row.Kind == CodeDiffKind.Removed ? '-' : ' ')
                    .Append(row.Text).Append("\r\n");
            return diff.ToString();
        }

        public static CodeDiffLine[] BuildRows(string before, string after)
        {
            var oldLines = CodeRollback.Lines(before); var newLines = CodeRollback.Lines(after);
            var hunks = CodeRollback.Hunks(before, after);
            var rows = new System.Collections.Generic.List<CodeDiffLine>();
            foreach (var hunk in hunks)
            {
                rows.Add(new CodeDiffLine { Kind = CodeDiffKind.Notice, Text = "Bloc " + (hunk.Index + 1) });
                for (int i = Math.Max(0, hunk.BeforeStart - 3); i < hunk.BeforeStart; i++)
                    rows.Add(new CodeDiffLine { Kind = CodeDiffKind.Context, OldLine = i + 1, NewLine = hunk.AfterStart - hunk.BeforeStart + i + 1, Text = oldLines[i] });
                for (int i = 0; i < hunk.Before.Length; i++)
                    rows.Add(new CodeDiffLine { Kind = CodeDiffKind.Removed, OldLine = hunk.BeforeStart + i + 1, Text = hunk.Before[i] });
                for (int i = 0; i < hunk.After.Length; i++)
                    rows.Add(new CodeDiffLine { Kind = CodeDiffKind.Added, NewLine = hunk.AfterStart + i + 1, Text = hunk.After[i] });
                for (int i = hunk.AfterStart + hunk.After.Length; i < Math.Min(newLines.Length, hunk.AfterStart + hunk.After.Length + 3); i++)
                    rows.Add(new CodeDiffLine { Kind = CodeDiffKind.Context, OldLine = hunk.BeforeStart + hunk.Before.Length + i - hunk.AfterStart - hunk.After.Length + 1, NewLine = i + 1, Text = newLines[i] });
            }
            return rows.Count == 0 ? new[] { new CodeDiffLine { Kind = CodeDiffKind.Notice, Text = "Aucune différence de code." } } : rows.ToArray();
        }

        private static string[] Lines(string code)
        {
            if (string.IsNullOrEmpty(code)) return new string[0];
            string[] lines = code.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
            if (lines.Length > 0 && lines[lines.Length - 1].Length == 0)
                Array.Resize(ref lines, lines.Length - 1);
            return lines;
        }
    }
}
