using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text.RegularExpressions;
using Markdig;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;
using Markdig.Extensions.Tables;
using Markdig.Extensions.TaskLists;
namespace VBAi
{

    /// <summary>Renders provider Markdown in the native chat transcript, including code, tables, links, and VBA references.</summary>
    internal static class ChatNativeMarkdown
    {

        /// <summary>Parses the Markdown extensions supported by the native transcript renderer.</summary>
        private static readonly MarkdownPipeline Pipeline = new MarkdownPipelineBuilder().UsePipeTables().UseAutoLinks().UseEmphasisExtras().UseTaskLists().Build();

        /// <summary>Parses Markdown and renders its blocks and interactive references into the transcript view.</summary>
        /// <param name="view">Transcript control that receives the rendered content.</param>
        /// <param name="text">Markdown source returned by the chat provider.</param>
        /// <param name="refs">Recognized VBA references keyed by their visible token.</param>
        /// <param name="navigate">Callback invoked when the user activates a recognized VBA reference.</param>
        /// <param name="error">Callback that receives link activation errors.</param>
        internal static void Render(ChatTextContentView view, string text, IDictionary<string,VbeChatReference> refs, Action<VbeChatReference> navigate, Action<string> error)
        { Blocks(view, Markdown.Parse(text,Pipeline), refs,navigate,error); }

        /// <summary>Renders Markdown blocks, including nested lists, tables, quotes, and code, into the transcript view.</summary>
        /// <param name="view">chat text content view that supplies the view for this operation.</param>
        /// <param name="blocks">Parsed Markdown blocks to render.</param>
        /// <param name="refs">i dictionary&lt;string,vbe chat reference&gt; that supplies the refs for this operation.</param>
        /// <param name="navigate">action&lt;vbe chat reference&gt; that supplies the navigate for this operation.</param>
        /// <param name="error">action&lt;string&gt; that supplies the error for this operation.</param>
        private static void Blocks(ChatTextContentView view, ContainerBlock blocks, IDictionary<string,VbeChatReference> refs, Action<VbeChatReference> navigate, Action<string> error)
        {
            foreach (var block in blocks) {
                if (block is CodeBlock code) {
                    string source = code.Lines.ToString(); int start = view.content.TextLength;
                    foreach (var part in VbaSyntax.Parts(source)) view.Append(part.Text,"Consolas",9,FontStyle.Regular,VbaSyntax.Color(part.Kind));
                    view.actions.Add(new ChatTextContentView.TextAction { Start = start, Length = source.Length, Code = source });
                    view.Append("\n");
                } else if (block is ThematicBreakBlock) view.Append("--------\n");
                else if (block is ListBlock list) {
                    int number; if (!int.TryParse(list.OrderedStart,out number) || number < 1) number = 1;
                    foreach (ListItemBlock item in list) { view.Append(list.IsOrdered ? (number++) + ". " : "• "); Blocks(view,item,refs,navigate,error); }
                } else if (block is Table table) {
                    foreach (TableRow row in table) {
                        foreach (TableCell cell in row) { foreach (var item in cell) if (item is LeafBlock leafCell && leafCell.Inline != null) Inlines(view,leafCell.Inline,refs,navigate,error,row.IsHeader ? FontStyle.Bold : FontStyle.Regular,9.5f); view.Append("\t"); }
                        view.Append("\n");
                    }
                } else if (block is QuoteBlock quote) { view.Append("│ "); Blocks(view,quote,refs,navigate,error); }
                else if (block is ContainerBlock nested) Blocks(view,nested,refs,navigate,error);
                else if (block is LeafBlock leaf) {
                    float size = block is HeadingBlock heading ? Math.Max(11, 18-heading.Level) : 9.5f;
                    var style = block is HeadingBlock ? FontStyle.Bold : FontStyle.Regular;
                    if (leaf.Inline != null) Inlines(view,leaf.Inline,refs,navigate,error,style,size);
                    else view.Append(leaf.Lines.ToString());
                    view.Append("\n");
                }
            }
        }

        /// <summary>Renders inline Markdown and wires allowed links and recognized VBA references to their actions.</summary>
        /// <param name="view">chat text content view that supplies the view for this operation.</param>
        /// <param name="source">Parsed inline Markdown content to render.</param>
        /// <param name="refs">i dictionary&lt;string,vbe chat reference&gt; that supplies the refs for this operation.</param>
        /// <param name="navigate">action&lt;vbe chat reference&gt; that supplies the navigate for this operation.</param>
        /// <param name="error">action&lt;string&gt; that supplies the error for this operation.</param>
        /// <param name="style">font style that supplies the style for this operation.</param>
        /// <param name="size">float that supplies the size for this operation.</param>
        private static void Inlines(ChatTextContentView view, ContainerInline source, IDictionary<string,VbeChatReference> refs, Action<VbeChatReference> navigate, Action<string> error, FontStyle style, float size)
        {
            foreach (var inline in source) {
                if (inline is LiteralInline literal) {
                    foreach (string part in Regex.Split(literal.Content.ToString(), "([#@][\\p{L}\\p{N}_.:]+)")) {
                        int start = view.content.TextLength;
                        bool link = refs != null && refs.TryGetValue(part,out var _);
                        view.Append(part,"Segoe UI",size,style | (link ? FontStyle.Underline : FontStyle.Regular),link ? (Color?)Color.RoyalBlue : null);
                        if (link) { var reference = refs[part]; view.actions.Add(new ChatTextContentView.TextAction { Start = start, Length = part.Length, Invoke = () => navigate(reference) }); }
                    }
                } else if (inline is CodeInline code) view.Append(code.Content,"Consolas",size,style);
                else if (inline is EmphasisInline emphasis) Inlines(view,emphasis,refs,navigate,error,style | (emphasis.DelimiterChar == '~' ? FontStyle.Strikeout : emphasis.DelimiterCount >= 2 ? FontStyle.Bold : FontStyle.Italic),size);
                else if (inline is LinkInline link) {
                    int start = view.content.TextLength; Inlines(view,link,refs,navigate,error,style | FontStyle.Underline,size);
                    if (SafeLinks.Allowed(link.Url)) view.actions.Add(new ChatTextContentView.TextAction { Start = start, Length = view.content.TextLength-start, Invoke = () => { try { ChatMarkdown.OpenLink(link.Url); } catch (Exception ex) { error(ex.Message); } } });
                } else if (inline is AutolinkInline auto) {
                    int start = view.content.TextLength; view.Append(auto.Url,"Segoe UI",size,style | FontStyle.Underline,Color.RoyalBlue);
                    if (SafeLinks.Allowed(auto.Url)) view.actions.Add(new ChatTextContentView.TextAction { Start=start, Length=auto.Url.Length, Invoke=() => { try { ChatMarkdown.OpenLink(auto.Url); } catch (Exception ex) { error(ex.Message); } } });
                } else if (inline is LineBreakInline br) view.Append(br.IsHard ? "\n" : " ");
                else if (inline is TaskList task) view.Append(task.Checked ? "☑ " : "☐ ");
                else if (inline is HtmlInline html) view.Append(html.Tag);
                else if (inline is ContainerInline nested) Inlines(view,nested,refs,navigate,error,style,size);
            }
        }
    }
}
