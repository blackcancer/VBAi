namespace CodexVBE.Tests.Unit
{
    using System;
    using System.IO;
    using System.Security.Cryptography;
    using System.Text;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    [TestClass]
    [TestCategory("Unit")]
    public sealed partial class CodeFileEncodingTests
    {
        [TestMethod]
        public void InspectionDistinguishesAsciiUtf8BomAndAmbiguousAnsi()
        {
            var root = Path.Combine(Path.GetTempPath(), "CodexVBE-EncodingTests", Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(root);
            try
            {
                var reader = new VbeCodeNavigation(null, new VbeForms(null));
                var ascii = Path.Combine(root, "plain.bas");
                File.WriteAllBytes(ascii, Encoding.ASCII.GetBytes("Option Explicit\r\n"));
                dynamic plain = reader.InspectCodeFile(ascii);
                Assert.AreEqual("utf-8", (string)plain.DefaultEncoding);
                Assert.IsFalse((bool)plain.ExplicitEncodingRequired);
                Assert.IsTrue((bool)plain.StrictUtf8Valid);
                Assert.IsFalse((bool)plain.ContentIncluded);
                var utf8 = Path.Combine(root, "utf8.bas");
                File.WriteAllBytes(utf8, Combine(new byte[] { 0xEF, 0xBB, 0xBF }, Encoding.UTF8.GetBytes("' été\r\n")));
                dynamic withBom = reader.InspectCodeFile(utf8);
                Assert.AreEqual("utf-8", (string)withBom.Bom);
                Assert.AreEqual("utf-8", (string)withBom.DefaultEncoding);
                Assert.IsTrue((bool)withBom.ContainsNonAscii);
                Assert.IsFalse((bool)withBom.ExplicitEncodingRequired);
                var ansi = Path.Combine(root, "ansi.bas");
                File.WriteAllBytes(ansi, new byte[] { 0x27, 0x20, 0xE9, 0x0D, 0x0A });
                dynamic ambiguous = reader.InspectCodeFile(ansi);
                Assert.IsNull((object)ambiguous.Bom);
                Assert.IsNull((object)ambiguous.DefaultEncoding);
                Assert.IsFalse((bool)ambiguous.StrictUtf8Valid);
                Assert.IsTrue((bool)ambiguous.ExplicitEncodingRequired);
            }
            finally
            {
                Directory.Delete(root, true);
            }
        }

        [TestMethod]
        public void InspectionReportsUtf16AndRejectsUnsafeFileInputs()
        {
            var root = Path.Combine(Path.GetTempPath(), "CodexVBE-EncodingTests", Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(root);
            try
            {
                var reader = new VbeCodeNavigation(null, new VbeForms(null));
                Assert.ThrowsException<ArgumentException>(() => reader.InspectCodeFile("relative.bas"));
                Assert.ThrowsException<FileNotFoundException>(() => reader.InspectCodeFile(Path.Combine(root, "missing.bas")));
                var empty = Path.Combine(root, "empty.bas");
                File.WriteAllBytes(empty, new byte[0]);
                Assert.ThrowsException<ArgumentException>(() => reader.InspectCodeFile(empty));
                var oversized = Path.Combine(root, "oversized.bas");
                File.WriteAllBytes(oversized, new byte[256 * 1024 + 1]);
                Assert.ThrowsException<ArgumentException>(() => reader.InspectCodeFile(oversized));
                var utf16 = Path.Combine(root, "utf16.bas");
                File.WriteAllBytes(utf16, Combine(new byte[] { 0xFF, 0xFE }, Encoding.Unicode.GetBytes("' été")));
                dynamic unicode = reader.InspectCodeFile(utf16);
                Assert.AreEqual("utf-16le", (string)unicode.Bom);
                Assert.IsTrue((bool)unicode.ContainsNulByte);
                Assert.IsNull((object)unicode.StrictUtf8Valid);
                Assert.IsFalse((bool)unicode.ExplicitEncodingRequired);
            }
            finally
            {
                Directory.Delete(root, true);
            }
        }

        [TestMethod]
        public void InsertionRequiresExplicitEncodingAndRollsBackCorruptedUnicode()
        {
            var root = Path.Combine(Path.GetTempPath(), "CodexVBE-EncodingTests", Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(root);
            try
            {
                var source = Path.Combine(root, "accent.bas");
                File.WriteAllBytes(source, Encoding.GetEncoding(1252).GetBytes("' été"));
                var module = new VbeSessionContractTests.FakeModule();
                var project = new VbeSessionContractTests.FakeProject
                {
                    Name = "Projet",
                    Mode = 2
                };
                project.VBComponents.Add(new VbeSessionContractTests.FakeComponent { Name = "Module1", Type = 1, CodeModule = module });
                var host = new VbeSessionContractTests.FakeVbe();
                host.VBProjects.Add(project);
                var navigation = new VbeCodeNavigation(host, new VbeForms(host));
                var request = new Request
                {
                    Project = "Projet",
                    Module = "Module1",
                    Path = source,
                    StartLine = 1,
                    ExpectedSha256 = Hash(string.Empty)
                };
                Assert.ThrowsException<ArgumentException>(() => navigation.InsertCodeFile(request));
                request.SourceEncoding = "utf-8";
                Assert.ThrowsException<ArgumentException>(() => navigation.InsertCodeFile(request));
                request.SourceEncoding = "windows-1252";
                module.CorruptNonAsciiOnInsert = true;
                Assert.ThrowsException<InvalidOperationException>(() => navigation.InsertCodeFile(request));
                Assert.AreEqual(0, module.CountOfLines, "Unicode corruption must roll back all inserted lines.");
                module.CorruptNonAsciiOnInsert = false;
                dynamic inserted = navigation.InsertCodeFile(request);
                Assert.AreEqual("windows-1252", (string)inserted.SourceEncoding);
                Assert.AreEqual(1, (int)inserted.InsertedLineCount);
                StringAssert.Contains(module.Code, "été");
                Assert.IsFalse((bool)inserted.CompilationVerified);
                Assert.ThrowsException<InvalidOperationException>(() => navigation.InsertCodeFile(request));
            }
            finally
            {
                Directory.Delete(root, true);
            }
        }
    }
}

namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Security.Cryptography;
    using System.Text;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    [TestClass]
    [TestCategory("Unit")]
    public sealed partial class VbeCodeNavigationCoverageTests
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
            dynamic exact = navigation.Find(new Request { Project = "Projet", Module = "module1", Query = "Alpha", MatchCase = true, WholeWord = true });
            Assert.AreEqual(1, (int)exact.Matches.Count);
            Assert.AreEqual(7, (int)exact.Matches[0].StartColumn);
        }

        [TestMethod]
        public void FindWildcardSearchRespectsWholeWordAndEscapedLiterals()
        {
            var navigation = CreateNavigation("foo_1 foo-2 xfoo_3\r\n[a] [b]", "unused");
            dynamic pattern = navigation.Find(new Request { Project = "Projet", Module = "Module1", Query = "foo?1", PatternSearch = true, WholeWord = true });
            Assert.AreEqual(1, (int)pattern.Matches.Count);
            dynamic embedded = navigation.Find(new Request { Project = "Projet", Module = "Module1", Query = "foo?3", PatternSearch = true, WholeWord = true });
            Assert.AreEqual(0, (int)embedded.Matches.Count);
            dynamic literal = navigation.Find(new Request { Project = "Projet", Module = "Module1", Query = "[?]", PatternSearch = true });
            Assert.AreEqual(2, (int)literal.Matches.Count);
            Assert.AreEqual(2, (int)literal.Matches[0].StartLine);
        }

        [TestMethod]
        public void FindRejectsInvalidQueriesAndUnknownModule()
        {
            var navigation = CreateNavigation("abc", "def");
            Assert.ThrowsException<ArgumentException>(() => navigation.Find(new Request { Project = "Projet", Query = "" }));
            Assert.ThrowsException<ArgumentException>(() => navigation.Find(new Request { Project = "Projet", Query = new string ('x', 201) }));
            Assert.ThrowsException<ArgumentException>(() => navigation.Find(new Request { Project = "Projet", Query = "***", PatternSearch = true }));
            Assert.ThrowsException<InvalidOperationException>(() => navigation.Find(new Request { Project = "Projet", Module = "Missing", Query = "abc" }));
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
            dynamic overlapping = navigation.Find(new Request { Project = "Projet", Module = "Module1", Query = "aa", MatchCase = true });
            Assert.AreEqual(3, (int)overlapping.Matches.Count);
            Assert.AreEqual(1, (int)overlapping.Matches[0].StartColumn);
            Assert.AreEqual(2, (int)overlapping.Matches[1].StartColumn);
            Assert.AreEqual(3, (int)overlapping.Matches[2].StartColumn);
            dynamic words = navigation.Find(new Request { Project = "Projet", Module = "Module1", Query = "alpha", WholeWord = true });
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
            dynamic exact = navigation.Find(new Request { Project = "Projet", Module = "Class1", Query = "A?pha", PatternSearch = true, MatchCase = true });
            Assert.AreEqual(1, (int)exact.Matches.Count);
            Assert.AreEqual(1, (int)exact.Matches[0].StartColumn);
            dynamic folded = navigation.Find(new Request { Project = "Projet", Module = "Class1", Query = "A?pha", PatternSearch = true, MatchCase = false });
            Assert.AreEqual(3, (int)folded.Matches.Count);
            Assert.AreEqual(13, (int)folded.Matches[2].StartColumn);
        }

        [TestMethod]
        public void ProcedureMutationRejectsBreakModeAndWrongComponentWithoutChangingCode()
        {
            var host = new VbeProcedureMutationTests.FakeVbe();
            var project = new VbeProcedureMutationTests.FakeProject
            {
                Name = "Projet",
                Mode = 1
            };
            var component = new VbeProcedureMutationTests.FakeComponent
            {
                Name = "Module1",
                Type = 1
            };
            var module = new VbeProcedureMutationTests.FakeModule(component, "Option Explicit");
            component.CodeModule = module;
            project.VBComponents.Add(component);
            host.VBProjects.Add(project);
            var navigation = new VbeCodeNavigation(host, new VbeForms(host));
            var request = new Request
            {
                Project = "Projet",
                Module = "Module1",
                Procedure = "Run",
                ProcKind = 0,
                Text = "Sub Run()\nEnd Sub",
                ExpectedSha256 = Hash(module.Code)
            };
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
            var request = new Request
            {
                Project = "Projet",
                Module = "Module1",
                Procedure = "Run",
                ProcKind = 0,
                ExpectedSha256 = Hash("stale")
            };
            var error = Assert.ThrowsException<InvalidOperationException>(() => navigation.SelectProcedure(request, null));
            StringAssert.Contains(error.Message, "module changed since it was read");
            Assert.AreEqual(0, request.StartLine);
        }

        [TestMethod]
        public void ProcedureCommandsRejectMalformedOrStaleInputBeforeEditing()
        {
            var navigation = CreateNavigation("Option Explicit", "");
            var request = new Request
            {
                Project = "Projet",
                Module = "Module1",
                Procedure = "Run",
                ProcKind = 0,
                Text = "Sub Run()\nEnd Sub",
                ExpectedSha256 = Hash("Option Explicit")
            };
            Assert.ThrowsException<ArgumentException>(() => navigation.CreateProcedure(new Request()));
            request.Text = "Sub Run()\nSub Other()\nEnd Sub";
            Assert.ThrowsException<ArgumentException>(() => navigation.CreateProcedure(request));
            request.Text = "Sub Run()\nEnd Sub";
            Assert.ThrowsException<InvalidOperationException>(() => navigation.CreateProcedure(request));
            Assert.ThrowsException<ArgumentException>(() => navigation.RemoveProcedure(new Request { Project = "Projet", Module = "Module1", Procedure = "bad name", ExpectedSha256 = request.ExpectedSha256 }));
            Assert.ThrowsException<ArgumentException>(() => navigation.SelectProcedure(new Request { Project = "Projet", Module = "Module1", Procedure = "Run" }, null));
        }

        [TestMethod]
        public void ProcedureValidationCoversEventAndPropertyDeclarations()
        {
            var navigation = CreateNavigation("Option Explicit", "");
            var currentHash = Hash("Option Explicit");
            Assert.ThrowsException<ArgumentException>(() => navigation.CreateEventProcedure(new Request { EventName = "Click", ObjectName = "Bad Name", ExpectedSha256 = currentHash, ExpectedTreeVersion = "tree" }));
            Assert.ThrowsException<ArgumentException>(() => navigation.CreateEventProcedure(new Request { EventName = "Click!", ObjectName = "Button1", ExpectedSha256 = currentHash, ExpectedTreeVersion = "tree" }));
            var property = new Request
            {
                Project = "Projet",
                Module = "Class1",
                Procedure = "Value",
                ProcKind = 3,
                ExpectedSha256 = Hash(""),
                Text = "Public Property Get Value() As String\nEnd Property"
            };
            Assert.ThrowsException<InvalidOperationException>(() => navigation.CreateProcedure(property));
            Assert.ThrowsException<InvalidOperationException>(() => navigation.ReplaceProcedure(property));
            Assert.ThrowsException<InvalidOperationException>(() => navigation.RemoveProcedure(property));
            property.Text = "Public Property Let Value(value As String)\nEnd Property";
            Assert.ThrowsException<ArgumentException>(() => navigation.CreateProcedure(property));
            property.ProcKind = 1;
            Assert.ThrowsException<InvalidOperationException>(() => navigation.CreateProcedure(property));
        }
    }
}

namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using System.Dynamic;
    using System.Linq;
    using System.Security.Cryptography;
    using System.Text;
    using System.Text.RegularExpressions;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    [TestClass]
    [TestCategory("Unit")]
    public sealed partial class VbeProcedureMutationTests
    {
        [TestMethod]
        public void ProjectSymbolsReturnsLiveDefinitionsAndPaginates()
        {
            var fixture = new Fixture("Option Explicit");
            fixture.Navigation.CreateProcedure(fixture.Request("Sub Run()\nDebug.Print 1\nEnd Sub"));
            dynamic all = fixture.Navigation.ProjectSymbols(new Request { Project = "Projet", Limit = 1 });
            Assert.AreEqual(2, (int)all.Total);
            Assert.IsTrue((bool)all.HasMore);
            Assert.AreEqual(0, (int)all.Errors.Count);
            dynamic found = fixture.Navigation.ProjectSymbols(new Request { Project = "Projet", Query = "run", WholeWord = true });
            Assert.AreEqual(1, (int)found.Total);
            Assert.AreEqual("Run", (string)found.Symbols[0].Name);
            Assert.AreEqual(3, (int)found.Symbols[0].Line);
            Assert.IsFalse(string.IsNullOrEmpty((string)found.Symbols[0].Sha256));
            dynamic exactCase = fixture.Navigation.ProjectSymbols(new Request { Project = "Projet", Query = "run", WholeWord = true, MatchCase = true });
            Assert.AreEqual(0, (int)exactCase.Total);
        }

        [TestMethod]
        public void CreateReplaceListAndRemoveProcedureRoundTrip()
        {
            var fixture = new Fixture("Option Explicit");
            var create = fixture.Request("Sub Run()\nDebug.Print 1\nEnd Sub");
            dynamic made = fixture.Navigation.CreateProcedure(create);
            Assert.AreEqual(3, (int)made.BodyLine);
            Assert.AreEqual("Run", (string)made.Procedure);
            Assert.IsFalse((bool)made.CompilationVerified);
            dynamic listed = fixture.Navigation.Procedures("Projet", "Module1");
            Assert.AreEqual(1, (int)listed.Procedures.Count);
            Assert.AreEqual("Run", (string)listed.Procedures[0].Name);
            create.ExpectedSha256 = (string)made.Sha256;
            create.Text = "Sub Run()\nDebug.Print 2\nEnd Sub";
            dynamic replaced = fixture.Navigation.ReplaceProcedure(create);
            Assert.IsTrue((bool)replaced.Changed);
            StringAssert.Contains(fixture.Module.Code, "Debug.Print 2");
            create.ExpectedSha256 = (string)replaced.Sha256;
            dynamic unchanged = fixture.Navigation.ReplaceProcedure(create);
            Assert.IsFalse((bool)unchanged.Changed);
            dynamic removed = fixture.Navigation.RemoveProcedure(create);
            Assert.AreEqual(3, (int)removed.RemovedLineCount);
            Assert.IsFalse(fixture.Module.Code.Contains("Run()"));
            Assert.AreEqual(0, (int)((dynamic)fixture.Navigation.Procedures("Projet", "Module1")).Procedures.Count);
        }

        [TestMethod]
        public void FunctionReplacementAndRemovalUseEndFunctionBoundary()
        {
            var fixture = new Fixture("Option Explicit");
            var request = fixture.Request("Public Function Run() As Long\nRun = 1\nEnd Function");
            dynamic created = fixture.Navigation.CreateProcedure(request);
            Assert.AreEqual("Run", (string)created.Procedure);
            request.ExpectedSha256 = (string)created.Sha256;
            request.Text = "Public Function Run() As Long\nRun = 2\nEnd Function";
            dynamic replaced = fixture.Navigation.ReplaceProcedure(request);
            Assert.IsTrue((bool)replaced.Changed);
            StringAssert.Contains(fixture.Module.Code, "Run = 2");
            request.ExpectedSha256 = (string)replaced.Sha256;
            dynamic removed = fixture.Navigation.RemoveProcedure(request);
            Assert.AreEqual(3, (int)removed.RemovedLineCount);
            Assert.IsFalse(fixture.Module.Code.Contains("Function Run"));
        }

        [TestMethod]
        public void StaleVersionAndDuplicateProcedureNeverChangeCode()
        {
            var fixture = new Fixture("Option Explicit");
            var request = fixture.Request("Sub Run()\nEnd Sub");
            request.ExpectedSha256 = Hash("stale");
            Assert.ThrowsException<InvalidOperationException>(() => fixture.Navigation.CreateProcedure(request));
            Assert.AreEqual("Option Explicit", fixture.Module.Code);
            request.ExpectedSha256 = Hash(fixture.Module.Code);
            fixture.Navigation.CreateProcedure(request);
            var before = fixture.Module.Code;
            request.ExpectedSha256 = Hash(before);
            Assert.ThrowsException<InvalidOperationException>(() => fixture.Navigation.CreateProcedure(request));
            Assert.AreEqual(before, fixture.Module.Code);
        }

        [TestMethod]
        public void FailedReplacementRestoresExactPreviousModuleText()
        {
            var fixture = new Fixture("Option Explicit\r\nSub Run()\r\nDebug.Print 1\r\nEnd Sub");
            var request = fixture.Request("Sub Run()\nDebug.Print 2\nEnd Sub");
            var before = fixture.Module.Code;
            fixture.Module.FailNextInsert = true;
            Assert.ThrowsException<InvalidOperationException>(() => fixture.Navigation.ReplaceProcedure(request));
            Assert.AreEqual(before, fixture.Module.Code);
        }

        [TestMethod]
        public void FailedRemovalRestoresExactPreviousModuleText()
        {
            var fixture = new Fixture("Option Explicit\r\nSub Run()\r\nEnd Sub");
            var request = fixture.Request(null);
            var before = fixture.Module.Code;
            fixture.Module.PretendStillPresentAfterDelete = true;
            Assert.ThrowsException<InvalidOperationException>(() => fixture.Navigation.RemoveProcedure(request));
            Assert.AreEqual(before, fixture.Module.Code);
        }

        [TestMethod]
        public void EventStubUsesCurrentFormTreeAndCurrentCodeVersion()
        {
            var fixture = new Fixture("", 3);
            fixture.Component.Designer.Controls.AddExisting("Button1");
            dynamic tree = fixture.Forms.Tree("Projet", "Module1");
            var request = new Request
            {
                Project = "Projet",
                Form = "Module1",
                ObjectName = "Button1",
                EventName = "Click",
                ExpectedTreeVersion = tree.TreeVersion,
                ExpectedSha256 = Hash("")
            };
            dynamic result = fixture.Navigation.CreateEventProcedure(request);
            Assert.AreEqual("Button1_Click", (string)result.Procedure);
            StringAssert.Contains(fixture.Module.Code, "Private Sub Button1_Click()");
            Assert.ThrowsException<InvalidOperationException>(() => fixture.Navigation.CreateEventProcedure(request));
        }

        [TestMethod]
        public void EventStubRejectsStaleTreeAndUnknownControlBeforeEditingCode()
        {
            var fixture = new Fixture("", 3);
            dynamic tree = fixture.Forms.Tree("Projet", "Module1");
            var request = new Request
            {
                Project = "Projet",
                Form = "Module1",
                ObjectName = "Missing",
                EventName = "Click",
                ExpectedTreeVersion = tree.TreeVersion,
                ExpectedSha256 = Hash("")
            };
            Assert.ThrowsException<InvalidOperationException>(() => fixture.Navigation.CreateEventProcedure(request));
            Assert.AreEqual("", fixture.Module.Code);
            request.ObjectName = "UserForm";
            request.ExpectedTreeVersion = "stale";
            Assert.ThrowsException<InvalidOperationException>(() => fixture.Navigation.CreateEventProcedure(request));
            Assert.AreEqual("", fixture.Module.Code);
        }
    }
}
