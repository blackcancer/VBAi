using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Integration
{
    /// <summary>Pins the native default root font and the declared fractional Frame font independently of import.</summary>
    internal static class UserFormQualificationFonts
    {
        internal const string OptIn = "VBAi_RUN_USERFORM_EXACT_FONT_QUALIFICATION";
        internal static bool Enabled => Environment.GetEnvironmentVariable(OptIn) == "1";

        /// <summary>Checks exact native scalars, including the fractional size; no rounding tolerance is accepted.</summary>
        internal static void RequireNative(IDictionary<string, object> fonts, string layout)
        {
            RequireOwner(fonts, "Form.Font", 8.25m);
            if (layout == "FrameMultiPage") RequireOwner(fonts, "Frame.Font", 8.27m);
        }

        private static void RequireOwner(IDictionary<string, object> fonts, string owner, decimal size)
        {
            Assert.AreEqual("Tahoma", Convert.ToString(fonts[owner + ".Name"]), owner + " face");
            Assert.AreEqual(size, Convert.ToDecimal(fonts[owner + ".Size"]), owner + " exact size");
            foreach (string member in new[] { "Bold", "Italic", "Underline", "Strikethrough" })
                Assert.AreEqual(false, Convert.ToBoolean(fonts[owner + "." + member]), owner + "." + member);
            Assert.AreEqual(0, Convert.ToInt32(fonts[owner + ".Charset"]), owner + " charset");
            Assert.AreEqual(400, Convert.ToInt32(fonts[owner + ".Weight"]), owner + " weight");
        }

        /// <summary>Accepts an omitted default root descriptor only with exact current native font evidence.</summary>
        internal static void RequireSnapshot(VbaGitSnapshot snapshot,
            Func<string, string, IDictionary<string, object>> readCurrentNativeFonts, string layout = null)
        {
            if (!Enabled) return;
            Assert.IsNotNull(readCurrentNativeFonts, "The current native form font reader is required.");
            foreach (var form in snapshot.Manifest.Components.Where(x => x.Type == 3))
            {
                var bindings = snapshot.FormFonts(form);
                Assert.IsNotNull(bindings, "Complete form resource grammar must expose its font bindings.");
                bool frameLayout = bindings.Any(x => x.Type == 14) || layout == "FrameMultiPage" ||
                    VbaGitSnapshot.Utf8.GetString(snapshot.Files[form.FileName]).Contains("LOCAL_GIT_LAYOUT_FrameMultiPage");
                var nativeFonts = readCurrentNativeFonts(form.Name, frameLayout ? "FrameMultiPage" : layout);
                Assert.IsNotNull(nativeFonts, form.Name + " current native font evidence is required.");
                RequireNative(nativeFonts, frameLayout ? "FrameMultiPage" : layout);
                var roots = bindings.Where(x => x.OwnerPath == "").ToArray();
                Assert.IsTrue(roots.Length <= 1 && roots.All(x => x.Type == 7),
                    form.Name + " has an unsupported or ambiguous root font plan.");
                if (roots.Length == 1)
                    CollectionAssert.AreEqual(Descriptor(8.25m), roots[0].Descriptor,
                        form.Name + " explicit default root descriptor");
                // MS-OFORMS may omit a default Form font. The current native getter above
                // then supplies the exact Tahoma 8.25/style/charset/weight proof.
                if (frameLayout) Assert.AreEqual(1, bindings.Count(x => x.Type == 14), "The declared Frame needs its explicit font binding.");
                foreach (var frame in bindings.Where(x => x.Type == 14))
                {
                    Assert.AreEqual("Controls/QualificationExtra", frame.OwnerPath, "Declared native Frame identity");
                    CollectionAssert.AreEqual(Descriptor(8.27m), frame.Descriptor, form.Name + " exact Frame descriptor");
                }
            }
        }

        /// <summary>Returns an independent fixed StdFont descriptor, with size expressed as currency units.</summary>
        internal static byte[] Descriptor(decimal size)
        {
            var value = new byte[] { 1, 0, 0, 0, 144, 1, 0, 0, 0, 0, 6, 84, 97, 104, 111, 109, 97 };
            Buffer.BlockCopy(BitConverter.GetBytes(checked((uint)(size * 10000m))), 0, value, 6, 4);
            return value;
        }
    }
}
