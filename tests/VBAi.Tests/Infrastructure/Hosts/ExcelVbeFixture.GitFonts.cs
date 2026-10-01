using System;
using System.Collections.Generic;
using System.Linq;

namespace VBAi.Tests.Integration
{
    internal sealed partial class ExcelVbeFixture
    {
        /// <summary>Reads actual root and nested Frame font values without assigning font members.</summary>
        internal IDictionary<string, object> ReadGitLayoutFonts(string form, string layout)
        {
            var result = new SortedDictionary<string, object>(StringComparer.Ordinal);
            WithGitLayoutDesigner(form, (component, designer) => {
                ReadGitFont(designer, "Form.Font", result);
                if (layout != "FrameMultiPage") return;
                object controls = null, frame = null;
                try
                {
                    controls = ((dynamic)designer).Controls;
                    frame = ((dynamic)controls).Item("QualificationExtra");
                    ReadGitFont(frame, "Frame.Font", result);
                }
                finally { Release(frame); Release(controls); }
            });
            return result;
        }

        /// <summary>Restores observed font values once on the owned synthetic form after a terminal import diagnostic.</summary>
        internal void RestoreGitLayoutFonts(string form, string layout, IDictionary<string, object> expected, bool assignOwner = false)
        {
            WithGitLayoutDesigner(form, (component, designer) => {
                AssignGitFontOnce(designer, "Form.Font", expected, assignOwner);
                if (layout != "FrameMultiPage") return;
                object controls = null, frame = null;
                try
                {
                    controls = ((dynamic)designer).Controls;
                    frame = ((dynamic)controls).Item("QualificationExtra");
                    AssignGitFontOnce(frame, "Frame.Font", expected, assignOwner);
                }
                finally { Release(frame); Release(controls); }
            });
        }

        /// <summary>Writes known font members once through guarded commands on Excel's actual owning STA.</summary>
        internal void RestoreGitLayoutFontsViaBridge(string form, string layout, IDictionary<string, object> expected)
        {
            object project = null;
            string projectName;
            try { project = ((dynamic)workbook).VBProject; projectName = Convert.ToString(((dynamic)project).Name); }
            finally { Release(project); }
            foreach (string member in new[] { "Name", "Size", "Bold", "Italic", "Underline", "Strikethrough" })
            {
                var tree = GitLayoutCommandData(Command(new { Command = "form_tree", Project = projectName, Form = form }));
                GitLayoutCommandData(Command(new { Command = "set_form_property", Project = projectName, Form = form,
                    ExpectedFormVersion = tree["TreeVersion"], Property = "Font." + member, Value = expected["Form.Font." + member] }));
            }
            if (layout != "FrameMultiPage") return;
            foreach (string member in new[] { "Name", "Size", "Bold", "Italic", "Underline", "Strikethrough" })
            {
                var tree = GitLayoutCommandData(Command(new { Command = "form_tree", Project = projectName, Form = form }));
                var frame = ((object[])tree["Controls"]).Select(VbeBridgeClient.Object)
                    .Single(node => Convert.ToString(node["Name"]) == "QualificationExtra");
                GitLayoutCommandData(Command(new { Command = "set_form_node_property", Project = projectName, Form = form,
                    ControlPath = frame["Path"], ExpectedTreeVersion = tree["TreeVersion"], Property = "Font." + member,
                    Value = expected["Frame.Font." + member] }));
            }
        }

        /// <summary>Assigns each declared native font member once; no setter is retried after an uncertain result.</summary>
        private static void AssignGitFontOnce(object owner, string prefix, IDictionary<string, object> expected, bool assignOwner)
        {
            object font = null;
            try
            {
                font = ((dynamic)owner).Font;
                ((dynamic)font).Name = Convert.ToString(expected[prefix + ".Name"]);
                ((dynamic)font).Size = Convert.ToDecimal(expected[prefix + ".Size"]);
                ((dynamic)font).Bold = Convert.ToBoolean(expected[prefix + ".Bold"]);
                ((dynamic)font).Italic = Convert.ToBoolean(expected[prefix + ".Italic"]);
                ((dynamic)font).Underline = Convert.ToBoolean(expected[prefix + ".Underline"]);
                ((dynamic)font).Strikethrough = Convert.ToBoolean(expected[prefix + ".Strikethrough"]);
                ((dynamic)font).Charset = Convert.ToInt16(expected[prefix + ".Charset"]);
                ((dynamic)font).Weight = Convert.ToInt16(expected[prefix + ".Weight"]);
                // An external StdFont can be marshalled by value. The owner setter
                // transfers the modified descriptor back into the actual designer.
                if (assignOwner) ((dynamic)owner).Font = font;
            }
            finally { Release(font); }
        }

        /// <summary>Reads the known native font members once and releases only a COM wrapper if one was returned.</summary>
        private static void ReadGitFont(object owner, string prefix, IDictionary<string, object> result)
        {
            object font = null;
            try
            {
                font = ((dynamic)owner).Font;
                result[prefix + ".Name"] = Convert.ToString(((dynamic)font).Name);
                result[prefix + ".Size"] = Convert.ToDouble(((dynamic)font).Size);
                result[prefix + ".Bold"] = Convert.ToBoolean(((dynamic)font).Bold);
                result[prefix + ".Italic"] = Convert.ToBoolean(((dynamic)font).Italic);
                result[prefix + ".Underline"] = Convert.ToBoolean(((dynamic)font).Underline);
                result[prefix + ".Strikethrough"] = Convert.ToBoolean(((dynamic)font).Strikethrough);
                var managed = font as System.Drawing.Font;
                if (managed != null)
                {
                    result[prefix + ".Charset"] = (int)managed.GdiCharSet;
                    result[prefix + ".Unit"] = managed.Unit.ToString();
                }
                else
                {
                    result[prefix + ".Charset"] = Convert.ToInt32(((dynamic)font).Charset);
                    result[prefix + ".Weight"] = Convert.ToInt32(((dynamic)font).Weight);
                }
            }
            finally { Release(font); }
        }
    }
}
