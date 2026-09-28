using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;
using Markdig;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;
using Markdig.Extensions.Tables;
using Markdig.Extensions.TaskLists;

namespace CodexVBE
{
    /// <summary>Convertit le Markdown du chat en contrôles WPF et relie références, liens et copie.</summary>
    internal static class ChatMarkdown
    {
        /// <summary>Action utilisée pour copier un bloc de code dans le presse-papiers.</summary>
        internal static Action<string> CopyText = Clipboard.SetText;
        /// <summary>Action utilisée pour ouvrir un lien HTTP(S) validé.</summary>
        internal static Action<string> OpenLink = SafeLinks.Open;
        /// <summary>Pipeline Markdig configuré pour tableaux, liens automatiques, emphase et cases à cocher.</summary>
        private static readonly MarkdownPipeline Pipeline = new MarkdownPipelineBuilder().UsePipeTables().UseAutoLinks().UseEmphasisExtras().UseTaskLists().Build();
        /// <summary>Crée un pinceau à partir d’une couleur ajustée par le thème.</summary>
        /// <param name="hex">Couleur au format hexadécimal.</param>
        /// <returns>Pinceau WPF correspondant à la couleur du thème.</returns>
        private static Brush Brush(string hex) { return new SolidColorBrush((Color)ColorConverter.ConvertFromString(UiTheme.Map(hex))); }
        /// <summary>Analyse le Markdown et construit un RichTextBox en lecture seule avec liens sûrs et références VBE navigables.</summary>
        /// <param name="text">Markdown à afficher.</param>
        /// <param name="references">Références reconnues dans le texte, éventuellement null.</param>
        /// <param name="navigate">Callback appelé quand une référence VBE est activée.</param>
        /// <param name="error">Callback qui reçoit les erreurs de copie ou d’ouverture de lien.</param>
        /// <returns>Contrôle WPF de conversation prêt à afficher.</returns>
        internal static RichTextBox Render(string text, IDictionary<string, VbeChatReference> references, Action<VbeChatReference> navigate, Action<string> error)
        {
            var document = new FlowDocument { PagePadding = new Thickness(0), FontFamily = new FontFamily("Segoe UI"), FontSize = 13,
                Foreground = Brush("#334155"), FlowDirection = UiText.Culture.TextInfo.IsRightToLeft ? FlowDirection.RightToLeft : FlowDirection.LeftToRight };
            AddBlocks(document.Blocks, Markdown.Parse(text ?? "", Pipeline), references, navigate, error);
            return new RichTextBox { Document = document, IsReadOnly = true, IsDocumentEnabled = true, BorderThickness = new Thickness(0),
                Background = Brushes.Transparent, Padding = new Thickness(0), VerticalScrollBarVisibility = ScrollBarVisibility.Disabled,
                HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled };
        }
        /// <summary>Rend récursivement les blocs de code, listes, tableaux, citations et paragraphes.</summary>
        /// <param name="target">Collection WPF qui reçoit les blocs produits.</param>
        /// <param name="blocks">Blocs Markdig à convertir.</param>
        /// <param name="refs">Références VBE associées au message.</param>
        /// <param name="navigate">Callback de navigation.</param>
        /// <param name="error">Callback de signalement des erreurs utilisateur.</param>
        private static void AddBlocks(BlockCollection target, ContainerBlock blocks, IDictionary<string, VbeChatReference> refs, Action<VbeChatReference> navigate, Action<string> error)
        {
            foreach (var block in blocks)
            {
                if (block is CodeBlock code)
                {
                    string source = code.Lines.ToString();
                    var panel = new StackPanel();
                    var copy = new Button { Content = UiText.Get("Copy code"), HorizontalAlignment = HorizontalAlignment.Right, Padding = new Thickness(8, 4, 8, 4), ToolTip = UiText.Get("Copy this code block to the clipboard.") };
                    copy.Click += (s, e) => { try { CopyText(source); } catch (Exception ex) { error(ex.Message); } };
                    panel.Children.Add(copy);
                    var doc = new FlowDocument { PagePadding = new Thickness(6), FontFamily = new FontFamily("Consolas"), FontSize = 12, FlowDirection = FlowDirection.LeftToRight };
                    var paragraph = new Paragraph { Margin = new Thickness(0) };
                    foreach (var part in VbaSyntax.Parts(source)) { var color = VbaSyntax.Color(part.Kind); paragraph.Inlines.Add(new Run(part.Text) { Foreground = new SolidColorBrush(Color.FromRgb(color.R, color.G, color.B)) }); }
                    doc.Blocks.Add(paragraph);
                    panel.Children.Add(new RichTextBox { Document = doc, IsReadOnly = true, BorderThickness = new Thickness(0), MaxHeight = 360,
                        Background = Brush("#F1F5F9"), VerticalScrollBarVisibility = ScrollBarVisibility.Auto, HorizontalScrollBarVisibility = ScrollBarVisibility.Auto });
                    target.Add(new BlockUIContainer(panel) { Margin = new Thickness(0, 6, 0, 10) });
                }
                else if (block is ListBlock list)
                {
                    var rendered = new System.Windows.Documents.List { MarkerStyle = list.IsOrdered ? TextMarkerStyle.Decimal : TextMarkerStyle.Disc, Margin = new Thickness(18, 3, 0, 6) };
                    int first; if (list.IsOrdered && int.TryParse(list.OrderedStart, out first) && first > 0) rendered.StartIndex = first;
                    foreach (ListItemBlock item in list) { var li = new ListItem(); AddBlocks(li.Blocks, item, refs, navigate, error); rendered.ListItems.Add(li); }
                    target.Add(rendered);
                }
                else if (block is Markdig.Extensions.Tables.Table table)
                {
                    var rendered = new System.Windows.Documents.Table { CellSpacing = 0, Margin = new Thickness(0, 6, 0, 10) };
                    var group = new TableRowGroup(); rendered.RowGroups.Add(group);
                    foreach (Markdig.Extensions.Tables.TableRow row in table)
                    {
                        var tr = new System.Windows.Documents.TableRow(); group.Rows.Add(tr);
                        foreach (Markdig.Extensions.Tables.TableCell cell in row)
                        {
                            var tc = new System.Windows.Documents.TableCell { BorderBrush = Brush("#E2E8F0"), BorderThickness = new Thickness(0.5), Padding = new Thickness(6), FontWeight = row.IsHeader ? FontWeights.SemiBold : FontWeights.Normal };
                            AddBlocks(tc.Blocks, cell, refs, navigate, error); tr.Cells.Add(tc);
                        }
                    }
                    target.Add(rendered);
                }
                else if (block is QuoteBlock quote)
                {
                    var section = new Section { BorderBrush = Brush("#64748B"), BorderThickness = new Thickness(3, 0, 0, 0), Padding = new Thickness(10, 2, 0, 2) };
                    AddBlocks(section.Blocks, quote, refs, navigate, error); target.Add(section);
                }
                else if (block is ThematicBreakBlock) target.Add(new Paragraph(new Run("────────")) { Foreground = Brush("#64748B") });
                else if (block is LeafBlock leaf)
                {
                    var paragraph = new Paragraph { Margin = new Thickness(0, 2, 0, 6) };
                    if (block is HeadingBlock heading) { paragraph.FontSize = Math.Max(14, 24 - heading.Level * 2); paragraph.FontWeight = FontWeights.SemiBold; }
                    if (leaf.Inline != null) AddInlines(paragraph.Inlines, leaf.Inline, refs, navigate, error);
                    else paragraph.Inlines.Add(new Run(leaf.Lines.ToString()));
                    target.Add(paragraph);
                }
            }
        }
        /// <summary>Rend récursivement les segments de texte, références, liens et styles inline.</summary>
        /// <param name="target">Collection WPF qui reçoit les segments produits.</param>
        /// <param name="source">Inlines Markdig à convertir.</param>
        /// <param name="refs">Références VBE associées au message.</param>
        /// <param name="navigate">Callback de navigation.</param>
        /// <param name="error">Callback de signalement des erreurs utilisateur.</param>
        private static void AddInlines(InlineCollection target, ContainerInline source, IDictionary<string, VbeChatReference> refs, Action<VbeChatReference> navigate, Action<string> error)
        {
            foreach (var inline in source)
            {
                if (inline is LiteralInline literal)
                {
                    foreach (string part in Regex.Split(literal.Content.ToString(), "([#@][\\p{L}\\p{N}_.:]+)"))
                    {
                        VbeChatReference reference;
                        if (refs != null && refs.TryGetValue(part, out reference)) { var link = new Hyperlink(new Run(part)); link.Click += (s, e) => navigate(reference); target.Add(link); }
                        else target.Add(new Run(part));
                    }
                }
                else if (inline is CodeInline code) target.Add(new Run(code.Content) { FontFamily = new FontFamily("Consolas"), Background = Brush("#F1F5F9") });
                else if (inline is LineBreakInline br) { if (br.IsHard) target.Add(new LineBreak()); else target.Add(new Run(" ")); }
                else if (inline is EmphasisInline emphasis)
                {
                    Span span = emphasis.DelimiterChar == '~' ? new Span { TextDecorations = TextDecorations.Strikethrough } : emphasis.DelimiterCount >= 2 ? (Span)new Bold() : new Italic();
                    AddInlines(span.Inlines, emphasis, refs, navigate, error); target.Add(span);
                }
                else if (inline is LinkInline link)
                {
                    var span = new Span(); AddInlines(span.Inlines, link, refs, navigate, error);
                    if (SafeLinks.Allowed(link.Url)) { var rendered = new Hyperlink(span) { ToolTip = link.Url }; rendered.Click += (s, e) => { try { OpenLink(link.Url); } catch (Exception ex) { error(ex.Message); } }; target.Add(rendered); }
                    else target.Add(span);
                }
                else if (inline is AutolinkInline auto)
                {
                    var autoLink = new Hyperlink(new Run(auto.Url)); autoLink.Click += (s, e) => { try { OpenLink(auto.Url); } catch (Exception ex) { error(ex.Message); } }; target.Add(autoLink);
                }
                else if (inline is TaskList task) target.Add(new Run(task.Checked ? "☑ " : "☐ "));
                else if (inline is HtmlInline html) target.Add(new Run(html.Tag));
                else if (inline is ContainerInline container) AddInlines(target, container, refs, navigate, error);
            }
        }
    }
}
