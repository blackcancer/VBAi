using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;
using System.Linq;

namespace VBAi.Tests.Integration
{
    internal sealed partial class OfficeVbeFixture
    {
        /// <summary>Maps the only selected saved Word project to its visible chat label before any UI action.</summary>
        internal string RequireWordChatGitScope(ExcelVbeFixture.EmbeddedGitScope scope)
        {
            Assert.AreEqual("Word", Kind);
            RequireUsableOwnedHost();
            RequireWordEmbeddedOwner(scope);
            RequireOwnedDocument();
            Assert.AreEqual(DocumentPath, scope.Path, true);
            Assert.AreEqual(DocumentPath, Data("debug_state", "Project", DocumentPath)["SelectedHostPath"]);
            var rows = Items("list_projects");
            var exact = rows.Where(row => string.Equals(VbeProjectHostPath.FromFields(row), DocumentPath,
                StringComparison.OrdinalIgnoreCase)).ToArray();
            Assert.AreEqual(1, exact.Length, "The owned Word document must have one canonical bridge project.");
            string label = Convert.ToString(exact[0]["Name"]) + " · " + Path.GetFileName(DocumentPath);
            Assert.AreEqual(1, rows.Count(row => string.Equals(Convert.ToString(row["Name"]) + " · " +
                (string.IsNullOrWhiteSpace(VbeProjectHostPath.FromFields(row)) ? "" :
                    Path.GetFileName(VbeProjectHostPath.FromFields(row))), label, StringComparison.Ordinal)),
                "A visible chat label cannot select an ambiguous Word document.");
            return label;
        }
    }
}
