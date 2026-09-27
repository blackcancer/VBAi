using System.Collections.Generic;
using System.Drawing;
using System.Text.RegularExpressions;
namespace CodexVBE
{
    internal static class VbaSyntax
    {
        internal sealed class Part { internal string Text, Kind; }
        private static readonly Regex Tokens = new Regex("(?<comment>'[^\\r\\n]*|\\bRem[ \\t][^\\r\\n]*)|(?<str>\"(?:\"\"|[^\"])*\")|(?<keyword>\\b(?:Sub|Function|Property|Get|Let|Set|End|If|Then|Else|ElseIf|For|Each|Next|While|Wend|Do|Loop|Until|Select|Case|Dim|As|Public|Private|Friend|Static|Const|ByVal|ByRef|Optional|Return|Exit|On|Error|Resume|GoTo|With|Nothing|True|False|New|Not|And|Or|Option|Explicit|Attribute|Integer|Long|String|Boolean|Variant|Object|Double)\\b)", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        internal static IEnumerable<Part> Parts(string source)
        {
            int end = 0;
            foreach (Match match in Tokens.Matches(source))
            {
                if (match.Index > end) yield return new Part { Text = source.Substring(end, match.Index - end), Kind = "plain" };
                yield return new Part { Text = match.Value, Kind = match.Groups["comment"].Success ? "comment" : match.Groups["str"].Success ? "string" : "keyword" };
                end = match.Index + match.Length;
            }
            if (end < source.Length) yield return new Part { Text = source.Substring(end), Kind = "plain" };
        }
        internal static Color Color(string kind) { return kind == "comment" ? (UiTheme.Dark ? System.Drawing.Color.LightGreen : System.Drawing.Color.ForestGreen) : kind == "string" ? (UiTheme.Dark ? System.Drawing.Color.SandyBrown : System.Drawing.Color.Brown) : kind == "keyword" ? (UiTheme.Dark ? System.Drawing.Color.LightSkyBlue : System.Drawing.Color.RoyalBlue) : UiTheme.Foreground; }
    }
}
