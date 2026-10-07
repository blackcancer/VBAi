using System;
using System.IO;
using System.Linq;
using System.Text;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class VbaRuntimeLicenseTests
    {
        /// <summary>Checks that a copied VBA module carries the entire authoritative grant, not just a license label.</summary>
        [TestMethod]
        public void CommentsPreserveTheCompleteEmbeddedMitGrantWithoutExecutableLines()
        {
            string expected;
            using (var stream = typeof(VbaRuntimeLicense).Assembly.GetManifestResourceStream("VBAi.Licenses.MIT.txt"))
            using (var reader = new StreamReader(stream, Encoding.UTF8, true))
                expected = reader.ReadToEnd().Replace("\r\n", "\n").TrimEnd('\n');

            string comments = VbaRuntimeLicense.Comments;
            StringAssert.StartsWith(comments, "' SPDX-License-Identifier: MIT\r\n");
            string notice = comments.Substring(comments.IndexOf("\r\n", StringComparison.Ordinal) + 2).TrimEnd('\r', '\n');
            var lines = notice.Split(new[] { "\r\n" }, StringSplitOptions.None);
            foreach (string line in lines)
            {
                Assert.IsTrue(line == "'" || line.StartsWith("' ", StringComparison.Ordinal));
                Assert.AreEqual(line.TrimEnd(' ', '\t'), line, "License comments must survive native VBE normalization.");
            }
            Assert.AreEqual(expected, string.Join("\n", lines.Select(line => line == "'" ? "" : line.Substring(2))));
            StringAssert.Contains(notice, "Copyright (c) 2026 VBAi contributors");
            StringAssert.Contains(notice, "THE SOFTWARE IS PROVIDED");
        }
    }
}
