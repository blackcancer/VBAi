// SPDX-License-Identifier: MPL-2.0
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Blackcancer.ChatUI.Presentation
{
    /// <summary>Detached table text. Formatting, units and unknown values retain their captured spelling.</summary>
    public sealed class ChatTableSnapshot
    {
        internal ChatTableSnapshot(string source, string[][] rows, int headerRowIndex, IEnumerable<ChatMarkdownTableAlignment> alignments)
        {
            SourceText = source ?? "";
            Rows = Array.AsReadOnly((rows ?? new string[0][]).Select(row =>
                (IReadOnlyList<string>)Array.AsReadOnly((row ?? new string[0]).Select(cell => cell ?? "").ToArray())).ToArray());
            HeaderRowIndex = headerRowIndex >= 0 && headerRowIndex < Rows.Count ? headerRowIndex : -1;
            ColumnAlignments = Array.AsReadOnly((alignments ?? Enumerable.Empty<ChatMarkdownTableAlignment>()).ToArray());
            ColumnCount = Math.Max(ColumnAlignments.Count, Rows.Select(row => row.Count).DefaultIfEmpty(0).Max());
        }

        public string SourceText { get; }
        public IReadOnlyList<IReadOnlyList<string>> Rows { get; }
        public IReadOnlyList<ChatMarkdownTableAlignment> ColumnAlignments { get; }
        public int HeaderRowIndex { get; }
        public int ColumnCount { get; }
        public int DataRowCount => Rows.Count - (HeaderRowIndex >= 0 ? 1 : 0);

        public string GetDataCell(int row, int column)
        {
            if (row < 0 || row >= DataRowCount || column < 0 || column >= ColumnCount) return "";
            int sourceRow = HeaderRowIndex >= 0 && row >= HeaderRowIndex ? row + 1 : row;
            return column < Rows[sourceRow].Count ? Rows[sourceRow][column] : "";
        }

        public string GetHeader(int column)
        {
            if (HeaderRowIndex < 0) return (column + 1).ToString(System.Globalization.CultureInfo.CurrentCulture);
            return column < Rows[HeaderRowIndex].Count ? Rows[HeaderRowIndex][column] : "";
        }

        public string ToTsv()
        {
            var result = new StringBuilder();
            for (int row = 0; row < Rows.Count; row++)
            {
                if (row > 0) result.Append(Environment.NewLine);
                for (int column = 0; column < ColumnCount; column++)
                {
                    if (column > 0) result.Append('\t');
                    string text = column < Rows[row].Count ? Rows[row][column] : "";
                    if (text.IndexOfAny(new[] { '\t', '\r', '\n', '"' }) >= 0)
                        result.Append('"').Append(text.Replace("\"", "\"\"")).Append('"');
                    else result.Append(text);
                }
            }
            return result.ToString();
        }
    }
}
