// SPDX-License-Identifier: MPL-2.0
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Drawing;

namespace Blackcancer.ChatUI.Presentation
{
    public enum ChatMarkdownBlockKind { Paragraph, Heading, Quote, ListItem, Code, Table, Rule }
    public enum ChatMarkdownTableAlignment { Unspecified, Left, Center, Right }

    /// <summary>Immutable host-independent transcript content, retaining its original Markdown.</summary>
    public sealed class ChatMarkdownDocument
    {
        internal ChatMarkdownDocument(string source, IList<ChatMarkdownBlock> blocks)
        {
            Source = source;
            Blocks = Freeze(blocks);
        }

        public string Source { get; }
        public IReadOnlyList<ChatMarkdownBlock> Blocks { get; }

        internal static IReadOnlyList<T> Freeze<T>(IList<T> items)
        {
            return new ReadOnlyCollection<T>(new List<T>(items));
        }
    }

    /// <summary>A renderable block; structural list and quote metadata survives flattening.</summary>
    public sealed class ChatMarkdownBlock
    {
        internal ChatMarkdownBlock(ChatMarkdownBlockKind kind, string source, IList<ChatMarkdownSpan> spans,
            string codeText = "", string language = "", int headingLevel = 0, int quoteDepth = 0,
            int listLevel = 0, string listPrefix = "", bool isTask = false, bool isChecked = false,
            IList<ChatMarkdownTableRow> tableRows = null, IList<ChatMarkdownTableColumn> tableColumns = null)
        {
            Kind = kind;
            Source = source;
            Spans = ChatMarkdownDocument.Freeze(spans);
            CodeText = codeText;
            Language = language;
            HeadingLevel = headingLevel;
            QuoteDepth = quoteDepth;
            ListLevel = listLevel;
            ListPrefix = listPrefix;
            IsTask = isTask;
            IsChecked = isChecked;
            TableRows = ChatMarkdownDocument.Freeze(tableRows ?? new List<ChatMarkdownTableRow>());
            TableColumns = ChatMarkdownDocument.Freeze(tableColumns ?? new List<ChatMarkdownTableColumn>());
        }

        public ChatMarkdownBlockKind Kind { get; }
        /// <summary>Original Markdown slice, including code fences where present.</summary>
        public string Source { get; }
        public IReadOnlyList<ChatMarkdownSpan> Spans { get; }
        /// <summary>Parsed code content, with indentation/fences removed according to Markdown rules.</summary>
        public string CodeText { get; }
        public string Language { get; }
        public int HeadingLevel { get; }
        public int QuoteDepth { get; }
        public int ListLevel { get; }
        public string ListPrefix { get; }
        public bool IsTask { get; }
        public bool IsChecked { get; }
        public IReadOnlyList<ChatMarkdownTableRow> TableRows { get; }
        public IReadOnlyList<ChatMarkdownTableColumn> TableColumns { get; }
    }

    public sealed class ChatMarkdownSpan
    {
        internal ChatMarkdownSpan(string text, FontStyle fontStyle, bool isCode = false, string href = null)
        {
            Text = text;
            FontStyle = fontStyle;
            IsCode = isCode;
            Href = href;
        }

        public string Text { get; }
        public FontStyle FontStyle { get; }
        public bool Bold => (FontStyle & FontStyle.Bold) != 0;
        public bool Italic => (FontStyle & FontStyle.Italic) != 0;
        public bool Strike => (FontStyle & FontStyle.Strikeout) != 0;
        public bool IsCode { get; }
        /// <summary>Inert link metadata. Navigation and URL policy belong to the rendering services.</summary>
        public string Href { get; }
    }

    public sealed class ChatMarkdownTableRow
    {
        internal ChatMarkdownTableRow(bool isHeader, IList<ChatMarkdownTableCell> cells)
        {
            IsHeader = isHeader;
            Cells = ChatMarkdownDocument.Freeze(cells);
        }

        public bool IsHeader { get; }
        public IReadOnlyList<ChatMarkdownTableCell> Cells { get; }
    }

    public sealed class ChatMarkdownTableCell
    {
        internal ChatMarkdownTableCell(string source, IList<ChatMarkdownSpan> spans, int columnIndex, int columnSpan)
        {
            Source = source;
            Spans = ChatMarkdownDocument.Freeze(spans);
            ColumnIndex = columnIndex;
            ColumnSpan = columnSpan;
        }

        public string Source { get; }
        public IReadOnlyList<ChatMarkdownSpan> Spans { get; }
        public int ColumnIndex { get; }
        public int ColumnSpan { get; }
    }

    public sealed class ChatMarkdownTableColumn
    {
        internal ChatMarkdownTableColumn(ChatMarkdownTableAlignment alignment) { Alignment = alignment; }
        public ChatMarkdownTableAlignment Alignment { get; }
    }
}
