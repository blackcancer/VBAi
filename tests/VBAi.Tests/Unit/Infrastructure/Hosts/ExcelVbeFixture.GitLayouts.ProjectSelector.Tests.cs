using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VBAi;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    /// <summary>Checks the saved workbook selector used by the Git layout and font bridge commands.</summary>
    [TestClass, TestCategory("Unit")]
    public sealed class ExcelVbeFixtureGitLayoutsProjectSelectorTests
    {
        public sealed class Project
        {
            public string Name { get; set; }
            public string FileName { get; set; }
        }

        public sealed class Vbe
        {
            public List<Project> VBProjects { get; } = new List<Project>();
        }

        [TestMethod]
        public void ExactOwnedHostPathSelectsOnlyItsProjectWhenNamesCollide()
        {
            string directory = Path.Combine(Path.GetTempPath(), "vbai-project-scope");
            string firstPath = Path.Combine(directory, "first.xlsm");
            string ownedPath = Path.Combine(directory, "owned.xlsm");
            var first = new Project { Name = "VBAProject", FileName = firstPath };
            var owned = new Project { Name = "VBAProject", FileName = ownedPath };
            var vbe = new Vbe(); vbe.VBProjects.Add(first); vbe.VBProjects.Add(owned);

            string selector = ExcelVbeFixture.RequireGitLayoutProjectSelector(
                ownedPath, directory, owned.FileName.ToUpperInvariant());
            Assert.AreEqual(Path.GetFullPath(ownedPath), selector, true);
            Assert.AreSame(owned, (object)VbeProjectResolver.Resolve(vbe, selector));
            Assert.ThrowsException<InvalidOperationException>(() => VbeProjectResolver.Resolve(vbe, "VBAProject"));
        }

        [TestMethod]
        public void UnsavedOrMismatchedWorkbookCannotDispatchAnyBridgeCommand()
        {
            string directory = Path.Combine(Path.GetTempPath(), "vbai-project-scope");
            string ownedPath = Path.Combine(directory, "owned.xlsm");
            string otherPath = Path.Combine(directory, "other.xlsm");
            foreach (var values in new[] {
                new[] { "Book1", "", "" },
                new[] { ownedPath, "", ownedPath },
                new[] { ownedPath, directory, "" },
                new[] { ownedPath, directory, otherPath },
                new[] { ownedPath, Path.Combine(directory, "other"), ownedPath },
                new[] { "C:owned.xlsm", directory, ownedPath },
                new[] { ownedPath, directory, "C:owned.xlsm" }
            })
                Assert.ThrowsException<InvalidOperationException>(() =>
                    ExcelVbeFixture.RequireGitLayoutProjectSelector(values[0], values[1], values[2]));
        }

        [TestMethod]
        public void ExactPathResolutionRefusesMissingOrDuplicateProjects()
        {
            string directory = Path.Combine(Path.GetTempPath(), "vbai-project-scope");
            string ownedPath = Path.Combine(directory, "owned.xlsm");
            string selector = ExcelVbeFixture.RequireGitLayoutProjectSelector(ownedPath, directory, ownedPath);
            var vbe = new Vbe();
            Assert.ThrowsException<InvalidOperationException>(() => VbeProjectResolver.Resolve(vbe, selector));
            vbe.VBProjects.Add(new Project { Name = "VBAProject", FileName = ownedPath });
            vbe.VBProjects.Add(new Project { Name = "VBAProject", FileName = ownedPath });
            Assert.ThrowsException<InvalidOperationException>(() => VbeProjectResolver.Resolve(vbe, selector));
        }
    }
}
