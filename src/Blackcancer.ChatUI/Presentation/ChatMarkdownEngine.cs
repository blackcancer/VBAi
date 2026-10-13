// SPDX-License-Identifier: MPL-2.0
// The pipeline and native Markdown traversal are adapted from VBAi's
// src/VBAi/Llm/Controls/Transcript/ChatNativeMarkdown.cs (MPL-2.0),
// local source baseline 213231485c3bba0251357357a09392b750b8c24f, and SolidAi's
// src/SolidAi.WinForms/Views/ChatNativeMarkdown.cs. Host actions and DTOs are excluded.
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using Markdig;
using Markdig.Extensions.Tables;
using Markdig.Extensions.TaskLists;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;

namespace Blackcancer.ChatUI.Presentation
{
    /// <summary>Parses native transcript Markdown without executing HTML or host actions.</summary>
    public static class ChatMarkdownEngine
    {
        private static readonly MarkdownPipeline Pipeline = new MarkdownPipelineBuilder()
            .UsePipeTables().UseAutoLinks().UseEmphasisExtras().UseTaskLists().Build();

        public static ChatMarkdownDocument Parse(string source)
        {
            source = source ?? string.Empty;
            var blocks = new List<ChatMarkdownBlock>();
            ReadBlocks(Markdown.Parse(source, Pipeline), source, blocks, 0, null);
            return new ChatMarkdownDocument(source, blocks);
        }

        private sealed class ListContext
        {
            internal int Level;
            internal string Prefix;
            internal bool PrefixUsed;
        }

        private static void ReadBlocks(ContainerBlock container, string source, List<ChatMarkdownBlock> output,
            int quoteDepth, ListContext listContext)
        {
            foreach (var block in container)
            {
                if (block is QuoteBlock quote)
                {
                    ReadBlocks(quote, source, output, quoteDepth + 1, listContext);
                    continue;
                }
                if (block is ListBlock list)
                {
                    int number;
                    if (!int.TryParse(list.OrderedStart, NumberStyles.Integer, CultureInfo.InvariantCulture, out number)) number = 1;
                    foreach (ListItemBlock item in list)
                    {
                        var context = new ListContext
                        {
                            Level = (listContext?.Level ?? 0) + 1,
                            Prefix = list.IsOrdered ? (number++).ToString(CultureInfo.InvariantCulture) + ". " : "• "
                        };
                        ReadBlocks(item, source, output, quoteDepth, context);
                    }
                    continue;
                }

                var spans = new List<ChatMarkdownSpan>();
                var kind = listContext != null ? ChatMarkdownBlockKind.ListItem :
                    quoteDepth > 0 ? ChatMarkdownBlockKind.Quote : ChatMarkdownBlockKind.Paragraph;
                var headingLevel = 0;
                string codeText = string.Empty, language = string.Empty;
                List<ChatMarkdownTableRow> rows = null;
                List<ChatMarkdownTableColumn> columns = null;
                bool isTask = false, isChecked = false;

                if (block is CodeBlock code)
                {
                    kind = ChatMarkdownBlockKind.Code;
                    codeText = code.Lines.ToString();
                    AddSpan(spans, codeText, FontStyle.Regular, true, null);
                    if (code is FencedCodeBlock fenced)
                    {
                        var info = (fenced.Info ?? string.Empty).Trim();
                        int separator = info.IndexOfAny(new[] { ' ', '\t' });
                        language = separator < 0 ? info : info.Substring(0, separator);
                    }
                }
                else if (block is Table table)
                {
                    kind = ChatMarkdownBlockKind.Table;
                    ReadTable(table, source, out rows, out columns);
                }
                else if (block is ThematicBreakBlock) kind = ChatMarkdownBlockKind.Rule;
                else if (block is ContainerBlock nested)
                {
                    ReadBlocks(nested, source, output, quoteDepth, listContext);
                    continue;
                }
                else if (block is LeafBlock leaf)
                {
                    var style = FontStyle.Regular;
                    if (block is HeadingBlock heading)
                    {
                        kind = ChatMarkdownBlockKind.Heading;
                        headingLevel = heading.Level;
                        style = FontStyle.Bold;
                    }
                    if (leaf.Inline != null)
                    {
                        ReadInlines(leaf.Inline, spans, style, null);
                        var task = FindTask(leaf.Inline);
                        isTask = task != null;
                        isChecked = task?.Checked ?? false;
                    }
                    else AddSpan(spans, leaf.Lines.ToString(), style, false, null);
                }
                else continue;

                var prefix = listContext != null && !listContext.PrefixUsed ? listContext.Prefix : string.Empty;
                if (listContext != null) listContext.PrefixUsed = true;
                output.Add(new ChatMarkdownBlock(kind, Slice(source, block), spans, codeText, language,
                    headingLevel, quoteDepth, listContext?.Level ?? 0, prefix, isTask, isChecked, rows, columns));
            }
        }

        private static void ReadTable(Table table, string source, out List<ChatMarkdownTableRow> rows,
            out List<ChatMarkdownTableColumn> columns)
        {
            rows = new List<ChatMarkdownTableRow>();
            columns = new List<ChatMarkdownTableColumn>();
            foreach (var definition in table.ColumnDefinitions)
            {
                var alignment = ChatMarkdownTableAlignment.Unspecified;
                if (definition.Alignment == TableColumnAlign.Left) alignment = ChatMarkdownTableAlignment.Left;
                else if (definition.Alignment == TableColumnAlign.Center) alignment = ChatMarkdownTableAlignment.Center;
                else if (definition.Alignment == TableColumnAlign.Right) alignment = ChatMarkdownTableAlignment.Right;
                columns.Add(new ChatMarkdownTableColumn(alignment));
            }
            foreach (TableRow row in table)
            {
                var cells = new List<ChatMarkdownTableCell>();
                int nextColumn = 0;
                foreach (TableCell cell in row)
                {
                    var spans = new List<ChatMarkdownSpan>();
                    ReadCell(cell, spans, row.IsHeader ? FontStyle.Bold : FontStyle.Regular);
                    // Pipe tables leave ColumnIndex at -1; explicit positions make cells usable by any grid.
                    int column = cell.ColumnIndex >= 0 ? cell.ColumnIndex : nextColumn;
                    int columnSpan = Math.Max(1, cell.ColumnSpan);
                    cells.Add(new ChatMarkdownTableCell(Slice(source, cell), spans, column, columnSpan));
                    nextColumn = column + columnSpan;
                }
                rows.Add(new ChatMarkdownTableRow(row.IsHeader, cells));
            }
            // Markdig 0.37 can retain one trailing definition beyond the cells.
            // Keep real empty cells; omit only definitions that have no cell in any row.
            int occupiedColumns = 0;
            foreach (var row in rows)
                foreach (var cell in row.Cells)
                    occupiedColumns = Math.Max(occupiedColumns, cell.ColumnIndex + cell.ColumnSpan);
            if (columns.Count > occupiedColumns) columns.RemoveRange(occupiedColumns, columns.Count - occupiedColumns);
        }

        private static void ReadCell(ContainerBlock cell, List<ChatMarkdownSpan> spans, FontStyle style)
        {
            foreach (var child in cell)
            {
                if (child is LeafBlock leaf)
                {
                    if (spans.Count > 0) AddSpan(spans, "\n", style, false, null);
                    if (leaf.Inline != null) ReadInlines(leaf.Inline, spans, style, null);
                    else AddSpan(spans, leaf.Lines.ToString(), style, false, null);
                }
                else if (child is ContainerBlock nested) ReadCell(nested, spans, style);
            }
        }

        private static void ReadInlines(ContainerInline source, List<ChatMarkdownSpan> output,
            FontStyle style, string href)
        {
            foreach (var inline in source)
            {
                if (inline is LiteralInline literal) AddSpan(output, literal.Content.ToString(), style, false, href);
                else if (inline is CodeInline code) AddSpan(output, code.Content, style, true, href);
                else if (inline is EmphasisInline emphasis)
                {
                    var emphasisStyle = emphasis.DelimiterChar == '~' ? FontStyle.Strikeout :
                        emphasis.DelimiterCount >= 2 ? FontStyle.Bold : FontStyle.Italic;
                    ReadInlines(emphasis, output, style | emphasisStyle, href);
                }
                else if (inline is LinkInline link)
                    ReadInlines(link, output, link.IsImage ? style : style | FontStyle.Underline,
                        link.IsImage ? href : link.Url);
                else if (inline is AutolinkInline auto)
                    AddSpan(output, auto.Url, style | FontStyle.Underline, false,
                        auto.IsEmail ? "mailto:" + auto.Url : auto.Url);
                else if (inline is LineBreakInline lineBreak) AddSpan(output, lineBreak.IsHard ? "\n" : " ", style, false, href);
                else if (inline is TaskList task) AddSpan(output, task.Checked ? "☑ " : "☐ ", style, false, href);
                else if (inline is HtmlInline html) AddSpan(output, html.Tag, style, false, href);
                else if (inline is HtmlEntityInline entity) AddSpan(output, entity.Transcoded.ToString(), style, false, href);
                else if (inline is ContainerInline nested) ReadInlines(nested, output, style, href);
            }
        }

        private static TaskList FindTask(ContainerInline source)
        {
            foreach (var inline in source)
            {
                if (inline is TaskList task) return task;
                if (inline is ContainerInline nested)
                {
                    var result = FindTask(nested);
                    if (result != null) return result;
                }
            }
            return null;
        }

        private static void AddSpan(List<ChatMarkdownSpan> output, string text, FontStyle style, bool isCode, string href)
        {
            if (!string.IsNullOrEmpty(text)) output.Add(new ChatMarkdownSpan(text, style, isCode, href));
        }

        private static string Slice(string source, MarkdownObject node)
        {
            var start = node.Span.Start;
            var end = node.Span.End;
            var code = node as FencedCodeBlock;
            if (code != null && code.ClosingFencedCharCount == 0 && start >= 0 && start < source.Length)
            {
                // An unfinished fence's Span.End still points into its opener.
                // Its line slices retain actual source positions, including partial streamed lines.
                int newline = source.IndexOf('\n', start);
                end = Math.Max(end, newline < 0 ? source.Length - 1 : newline - 1);
                for (int line = 0; line < code.Lines.Count; line++)
                    end = Math.Max(end, code.Lines.Lines[line].Slice.End);
            }
            return start >= 0 && end >= start && start < source.Length ?
                source.Substring(start, Math.Min(end, source.Length - 1) - start + 1) : string.Empty;
        }
    }
}
