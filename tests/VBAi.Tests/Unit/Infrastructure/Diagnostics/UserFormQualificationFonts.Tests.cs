using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    /// <summary>Checks that the native qualification oracle rejects the known fractional regression.</summary>
    [TestClass]
    public sealed partial class UserFormQualificationFontsTests
    {
        [TestMethod]
        public void RoundedFrameAndChangedDefaultRootCannotPass()
        {
            var values = Fonts();
            UserFormQualificationFonts.RequireNative(values, "FrameMultiPage");
            values["Frame.Font.Size"] = 8.25d;
            Assert.ThrowsException<AssertFailedException>(() => UserFormQualificationFonts.RequireNative(values, "FrameMultiPage"));
            values["Frame.Font.Size"] = 8.27d;
            values["Form.Font.Name"] = "Arial";
            Assert.ThrowsException<AssertFailedException>(() => UserFormQualificationFonts.RequireNative(values, "LabelButton"));
        }

        [TestMethod]
        public void ChangedDefaultStyleCharsetOrWeightCannotPass()
        {
            foreach (var change in new Dictionary<string, object>
            {
                ["Italic"] = true,
                ["Bold"] = true,
                ["Underline"] = true,
                ["Strikethrough"] = true,
                ["Charset"] = 1,
                ["Weight"] = 700
            })
            {
                var values = Fonts(); values["Form.Font." + change.Key] = change.Value;
                Assert.ThrowsException<AssertFailedException>(() => UserFormQualificationFonts.RequireNative(values, "LabelButton"));
            }
        }

        [TestMethod]
        public void RootAndFrameDescriptorsUseExactCurrencySizeWithoutAliasing()
        {
            byte[] root = UserFormQualificationFonts.Descriptor(8.25m);
            byte[] frame = UserFormQualificationFonts.Descriptor(8.27m);
            Assert.AreEqual(82500u, BitConverter.ToUInt32(root, 6));
            Assert.AreEqual(82700u, BitConverter.ToUInt32(frame, 6));
            Assert.AreEqual("Tahoma", System.Text.Encoding.ASCII.GetString(root, 11, root[10]));
            frame[11] = 0;
            Assert.AreEqual((byte)'T', UserFormQualificationFonts.Descriptor(8.27m)[11]);
        }

        private static Dictionary<string, object> Fonts()
        {
            var result = new Dictionary<string, object>();
            foreach (string owner in new[] { "Form.Font", "Frame.Font" })
            {
                result[owner + ".Name"] = "Tahoma";
                result[owner + ".Size"] = owner == "Frame.Font" ? 8.27d : 8.25d;
                result[owner + ".Charset"] = 0; result[owner + ".Weight"] = 400;
                foreach (string member in new[] { "Bold", "Italic", "Underline", "Strikethrough" }) result[owner + "." + member] = false;
            }
            return result;
        }
    }
}
