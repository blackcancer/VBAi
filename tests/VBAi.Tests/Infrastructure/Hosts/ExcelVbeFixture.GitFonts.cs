using System;
using System.Collections.Generic;

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
