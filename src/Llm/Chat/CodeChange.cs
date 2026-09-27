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
            string[] oldLines = Lines(before);
            string[] newLines = Lines(after);
            int prefix = 0;
            while (prefix < oldLines.Length && prefix < newLines.Length &&
                oldLines[prefix] == newLines[prefix]) prefix++;
            int suffix = 0;
            while (suffix < oldLines.Length - prefix && suffix < newLines.Length - prefix &&
                oldLines[oldLines.Length - suffix - 1] == newLines[newLines.Length - suffix - 1]) suffix++;
            if (prefix == oldLines.Length && prefix == newLines.Length)
                return new[] { new CodeDiffLine { Kind = CodeDiffKind.Notice, Text = "Aucune différence de code." } };
            var rows = new System.Collections.Generic.List<CodeDiffLine>();
            for (int line = Math.Max(0, prefix - 3); line < prefix; line++)
                rows.Add(new CodeDiffLine { Kind = CodeDiffKind.Context,
                    OldLine = line + 1, NewLine = line + 1, Text = oldLines[line] });
            for (int line = prefix; line < oldLines.Length - suffix; line++)
                rows.Add(new CodeDiffLine { Kind = CodeDiffKind.Removed,
                    OldLine = line + 1, Text = oldLines[line] });
            for (int line = prefix; line < newLines.Length - suffix; line++)
                rows.Add(new CodeDiffLine { Kind = CodeDiffKind.Added,
                    NewLine = line + 1, Text = newLines[line] });
            for (int line = 0; line < Math.Min(3, suffix); line++)
                rows.Add(new CodeDiffLine { Kind = CodeDiffKind.Context,
                    OldLine = oldLines.Length - suffix + line + 1,
                    NewLine = newLines.Length - suffix + line + 1,
                    Text = oldLines[oldLines.Length - suffix + line] });
            return rows.ToArray();
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
