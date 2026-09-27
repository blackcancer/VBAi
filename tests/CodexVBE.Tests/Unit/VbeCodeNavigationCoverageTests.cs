using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using CodexVBE;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Unit
{
    [TestClass]
    [TestCategory("Unit")]
    public sealed class VbeCodeNavigationCoverageTests
    {
        [TestMethod]
        public void FindSearchesEveryModuleAndHonorsCaseAndWholeWord()
        {
            var navigation = CreateNavigation("alpha Alpha alphabet\r\nALPHA", "Alpha");
            dynamic all = navigation.Find(new Request { Project = "Projet", Query = "alpha" });
            Assert.AreEqual(5, (int)all.Matches.Count);
            Assert.AreEqual(2, (int)all.SourceVersions.Count);
            Assert.AreEqual("Module1", (string)all.Matches[0].Module);
            Assert.AreEqual(1, (int)all.Matches[0].StartColumn);
            Assert.AreEqual(5, (int)all.Matches[0].EndColumn);

            dynamic words = navigation.Find(new Request { Project = "Projet", Query = "alpha", WholeWord = true });
            Assert.AreEqual(4, (int)words.Matches.Count);
            dynamic exact = navigation.Find(new Request { Project = "Projet", Module = "module1",
                Query = "Alpha", MatchCase = true, WholeWord = true });
            Assert.AreEqual(1, (int)exact.Matches.Count);
            Assert.AreEqual(7, (int)exact.Matches[0].StartColumn);
        }

        [TestMethod]
        public void FindWildcardSearchRespectsWholeWordAndEscapedLiterals()
        {
            var navigation = CreateNavigation("foo_1 foo-2 xfoo_3\r\n[a] [b]", "unused");
            dynamic pattern = navigation.Find(new Request { Project = "Projet", Module = "Module1",
                Query = "foo?1", PatternSearch = true, WholeWord = true });
            Assert.AreEqual(1, (int)pattern.Matches.Count);
            dynamic embedded = navigation.Find(new Request { Project = "Projet", Module = "Module1",
                Query = "foo?3", PatternSearch = true, WholeWord = true });
            Assert.AreEqual(0, (int)embedded.Matches.Count);
            dynamic literal = navigation.Find(new Request { Project = "Projet", Module = "Module1",
                Query = "[?]", PatternSearch = true });
            Assert.AreEqual(2, (int)literal.Matches.Count);
            Assert.AreEqual(2, (int)literal.Matches[0].StartLine);
        }

        [TestMethod]
        public void FindRejectsInvalidQueriesAndUnknownModule()
        {
            var navigation = CreateNavigation("abc", "def");
            Assert.ThrowsException<ArgumentException>(() => navigation.Find(new Request { Project = "Projet", Query = "" }));
            Assert.ThrowsException<ArgumentException>(() => navigation.Find(new Request { Project = "Projet",
                Query = new string('x', 201) }));
            Assert.ThrowsException<ArgumentException>(() => navigation.Find(new Request { Project = "Projet",
                Query = "***", PatternSearch = true }));
            Assert.ThrowsException<InvalidOperationException>(() => navigation.Find(new Request { Project = "Projet",
                Module = "Missing", Query = "abc" }));
        }

        [TestMethod]
        public void FindTruncatesLargeResultSetsAtTwoHundredMatches()
        {
            var navigation = CreateNavigation(string.Join("\r\n", Enumerable.Repeat("hit", 220)), "");
            dynamic result = navigation.Find(new Request { Project = "Projet", Module = "Module1", Query = "hit" });
            Assert.IsTrue((bool)result.Truncated);
            Assert.AreEqual(200, (int)result.Matches.Count);
            Assert.AreEqual(200, (int)result.Matches[199].StartLine);
        }

        [TestMethod]
        public void FindReportsOverlappingMatchesAndTreatsUnicodeAndUnderscoreAsWordCharacters()
        {
            var navigation = CreateNavigation("aaaa\r\néalpha alpha_ alpha é alpha", "");
            dynamic overlapping = navigation.Find(new Request { Project = "Projet", Module = "Module1",
                Query = "aa", MatchCase = true });
            Assert.AreEqual(3, (int)overlapping.Matches.Count);
            Assert.AreEqual(1, (int)overlapping.Matches[0].StartColumn);
            Assert.AreEqual(2, (int)overlapping.Matches[1].StartColumn);
            Assert.AreEqual(3, (int)overlapping.Matches[2].StartColumn);

            dynamic words = navigation.Find(new Request { Project = "Projet", Module = "Module1",
                Query = "alpha", WholeWord = true });
            Assert.AreEqual(2, (int)words.Matches.Count);
            Assert.AreEqual(2, (int)words.Matches[0].StartLine);
            Assert.AreEqual(15, (int)words.Matches[0].StartColumn);
            Assert.AreEqual(23, (int)words.Matches[1].StartColumn);
        }

        [TestMethod]
        public void FindIncludesEmptyModulesInSourceVersionsAndHonorsWildcardCase()
        {
            var navigation = CreateNavigation("", "Alpha alpha ALPHA");
            dynamic empty = navigation.Find(new Request { Project = "Projet", Query = "missing" });
            Assert.AreEqual(0, (int)empty.Matches.Count);
            Assert.AreEqual(2, (int)empty.SourceVersions.Count);
            Assert.IsFalse((bool)empty.Truncated);

            dynamic exact = navigation.Find(new Request { Project = "Projet", Module = "Class1",
                Query = "A?pha", PatternSearch = true, MatchCase = true });
            Assert.AreEqual(2, (int)exact.Matches.Count);
            Assert.AreEqual(1, (int)exact.Matches[0].StartColumn);
            Assert.AreEqual(13, (int)exact.Matches[1].StartColumn);
        }

        [TestMethod]
        public void ProcedureMutationRejectsBreakModeAndWrongComponentWithoutChangingCode()
        {
            var host = new VbeProcedureMutationTests.FakeVbe();
            var project = new VbeProcedureMutationTests.FakeProject { Name = "Projet", Mode = 1 };
            var component = new VbeProcedureMutationTests.FakeComponent { Name = "Module1", Type = 1 };
            var module = new VbeProcedureMutationTests.FakeModule(component, "Option Explicit");
            component.CodeModule = module;
            project.VBComponents.Add(component);
            host.VBProjects.Add(project);
            var navigation = new VbeCodeNavigation(host, new VbeForms(host));
            var request = new Request { Project = "Projet", Module = "Module1", Procedure = "Run",
                ProcKind = 0, Text = "Sub Run()\nEnd Sub", ExpectedSha256 = Hash(module.Code) };

            Assert.ThrowsException<InvalidOperationException>(() => navigation.CreateProcedure(request));
            Assert.ThrowsException<InvalidOperationException>(() => navigation.ReplaceProcedure(request));
            Assert.ThrowsException<InvalidOperationException>(() => navigation.RemoveProcedure(request));
            Assert.AreEqual("Option Explicit", module.Code);

            project.Mode = 2;
            component.Type = 3;
            Assert.ThrowsException<InvalidOperationException>(() => navigation.CreateProcedure(request));
            Assert.ThrowsException<InvalidOperationException>(() => navigation.ReplaceProcedure(request));
            Assert.ThrowsException<InvalidOperationException>(() => navigation.RemoveProcedure(request));
            Assert.AreEqual("Option Explicit", module.Code);
        }

        [TestMethod]
        public void SelectProcedureRejectsStaleHashBeforeLookingUpProcedureOrSelectingCode()
        {
            var navigation = CreateNavigation("Option Explicit", "");
            var request = new Request { Project = "Projet", Module = "Module1", Procedure = "Run",
                ProcKind = 0, ExpectedSha256 = Hash("stale") };
            var error = Assert.ThrowsException<InvalidOperationException>(() =>
                navigation.SelectProcedure(request, null));
            StringAssert.Contains(error.Message, "module changed since it was read");
            Assert.AreEqual(0, request.StartLine);
        }

        [TestMethod]
        public void ProcedureCommandsRejectMalformedOrStaleInputBeforeEditing()
        {
            var navigation = CreateNavigation("Option Explicit", "");
            var request = new Request { Project = "Projet", Module = "Module1", Procedure = "Run",
                ProcKind = 0, Text = "Sub Run()\nEnd Sub", ExpectedSha256 = Hash("Option Explicit") };
            Assert.ThrowsException<ArgumentException>(() => navigation.CreateProcedure(new Request()));
            request.Text = "Sub Run()\nSub Other()\nEnd Sub";
            Assert.ThrowsException<ArgumentException>(() => navigation.CreateProcedure(request));
            request.Text = "Sub Run()\nEnd Sub";
            Assert.ThrowsException<InvalidOperationException>(() => navigation.CreateProcedure(request));
            Assert.ThrowsException<ArgumentException>(() => navigation.RemoveProcedure(new Request {
                Project = "Projet", Module = "Module1", Procedure = "bad name", ExpectedSha256 = request.ExpectedSha256 }));
            Assert.ThrowsException<ArgumentException>(() => navigation.SelectProcedure(new Request {
                Project = "Projet", Module = "Module1", Procedure = "Run" }, null));
        }

        [TestMethod]
        public void ProcedureValidationCoversEventAndPropertyDeclarations()
        {
            var navigation = CreateNavigation("Option Explicit", "");
            var currentHash = Hash("Option Explicit");
            Assert.ThrowsException<ArgumentException>(() => navigation.CreateEventProcedure(new Request {
                EventName = "Click", ObjectName = "Bad Name", ExpectedSha256 = currentHash,
                ExpectedTreeVersion = "tree" }));
            Assert.ThrowsException<ArgumentException>(() => navigation.CreateEventProcedure(new Request {
                EventName = "Click!", ObjectName = "Button1", ExpectedSha256 = currentHash,
                ExpectedTreeVersion = "tree" }));

            var property = new Request { Project = "Projet", Module = "Class1", Procedure = "Value",
                ProcKind = 3, ExpectedSha256 = Hash(""), Text = "Public Property Get Value() As String\nEnd Property" };
            Assert.ThrowsException<InvalidOperationException>(() => navigation.CreateProcedure(property));
            Assert.ThrowsException<InvalidOperationException>(() => navigation.ReplaceProcedure(property));
            Assert.ThrowsException<InvalidOperationException>(() => navigation.RemoveProcedure(property));
            property.Text = "Public Property Let Value(value As String)\nEnd Property";
            Assert.ThrowsException<ArgumentException>(() => navigation.CreateProcedure(property));
            property.ProcKind = 1;
            Assert.ThrowsException<InvalidOperationException>(() => navigation.CreateProcedure(property));
        }

        private static VbeCodeNavigation CreateNavigation(string first, string second)
        {
            var host = new FakeVbe();
            var project = new FakeProject { Name = "Projet", Mode = 0 };
            project.VBComponents.Add(new FakeComponent { Name = "Module1", Type = 1, CodeModule = new FakeModule(first) });
            project.VBComponents.Add(new FakeComponent { Name = "Class1", Type = 2, CodeModule = new FakeModule(second) });
            host.VBProjects.Add(project);
            return new VbeCodeNavigation(host, new VbeForms(host));
        }

        private static string Hash(string source)
        {
            using (var sha = SHA256.Create())
                return BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(source)))
                    .Replace("-", "").ToLowerInvariant();
        }

        public sealed class FakeVbe { public List<FakeProject> VBProjects { get; } = new List<FakeProject>(); }
        public sealed class FakeProject
        {
            public string Name { get; set; }
            public int Mode { get; set; }
            public List<FakeComponent> VBComponents { get; } = new List<FakeComponent>();
        }
        public sealed class FakeComponent
        {
            public string Name { get; set; }
            public int Type { get; set; }
            public FakeModule CodeModule { get; set; }
        }
        public sealed class FakeModule
        {
            private readonly string[] lines;
            public FakeModule(string code)
            {
                lines = string.IsNullOrEmpty(code) ? new string[0] : code.Split(new[] { "\r\n" }, StringSplitOptions.None);
                Lines = new FakeLines(this);
            }
            public int CountOfLines => lines.Length;
            public FakeLines Lines { get; }
            public sealed class FakeLines
            {
                private readonly FakeModule module;
                public FakeLines(FakeModule module) { this.module = module; }
                public string this[int start, int count] => string.Join("\r\n", module.lines.Skip(start - 1).Take(count));
            }
        }
    }
}
