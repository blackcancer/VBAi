using Blackcancer.ChatUI.Views;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;

namespace VBAi.Tests.Infrastructure
{
    /// <summary>Inspects the actual shared Markdown controls and invokes their local handlers without input or focus automation.</summary>
    internal static class SharedMarkdownViewFixture
    {
        internal sealed class Link
        {
            internal RichTextBox Content;
            internal int Start, Length;
            internal Action Activate;
        }

        internal static IEnumerable<Control> Descendants(Control root)
        {
            foreach (Control child in root.Controls)
            {
                yield return child;
                foreach (Control descendant in Descendants(child)) yield return descendant;
            }
        }

        internal static ChatTextView[] TextViews(ChatTextContentView view)
            => Descendants(view.MarkdownView).OfType<ChatTextView>().ToArray();

        internal static ChatCodeBlockView[] CodeViews(ChatTextContentView view)
            => Descendants(view.MarkdownView).OfType<ChatCodeBlockView>().ToArray();

        internal static Link[] Links(ChatTextContentView view)
        {
            var links = new List<Link>();
            const BindingFlags flags = BindingFlags.Instance | BindingFlags.NonPublic;
            foreach (var text in TextViews(view))
            {
                var actions = (IEnumerable)typeof(ChatTextView).GetField("actions", flags).GetValue(text);
                foreach (var action in actions)
                {
                    var type = action.GetType();
                    int start = (int)type.GetField("Start", flags).GetValue(action);
                    links.Add(new Link
                    {
                        Content = text.Content,
                        Start = start,
                        Length = (int)type.GetField("Length", flags).GetValue(action),
                        Activate = () => typeof(ChatTextView).GetMethod("ActivateAt", flags).Invoke(text, new object[] { start })
                    });
                }
            }
            return links.ToArray();
        }

        internal static void CopyCode(ChatCodeBlockView code)
            => ((Button)Descendants(code).Single(control => control.Name == "copy")).PerformClick();
    }
}
