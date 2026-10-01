using System;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class VbaTestDiscoveryTests
    {
        [TestMethod]
        public void ExplicitAnnotationsDiscoverTestsAndAllFixtureRoles()
        {
            var catalog = Discover("Option Explicit\n'@TestModule\n" +
                "'@ModuleInitialize\nPublic Sub OpenFixture()\nEnd Sub\n" +
                "'@ModuleCleanup\nPublic Sub CloseFixture()\nEnd Sub\n" +
                "'@TestInitialize\nPublic Sub Prepare()\nEnd Sub\n" +
                "'@TestCleanup\nPublic Sub Restore()\nEnd Sub\n" +
                "'@TestMethod\nPublic Function CheckValue() As Boolean\nCheckValue = True\nEnd Function\n" +
                "'@TestMethod\nPublic Sub CheckAction()\nEnd Sub");
            Assert.AreEqual(2, catalog.Tests.Count());
            Assert.AreEqual(0, catalog.Diagnostics.Count);
            var module = catalog.Modules.Single();
            Assert.AreEqual("OpenFixture", module.ModuleInitialize.Procedure);
            Assert.AreEqual("CloseFixture", module.ModuleCleanup.Procedure);
            Assert.AreEqual("Prepare", module.TestInitialize.Procedure);
            Assert.AreEqual("Restore", module.TestCleanup.Procedure);
            CollectionAssert.AreEqual(new[] { "CheckAction", "CheckValue" }, catalog.Tests.Select(x => x.Procedure).ToArray());
            Assert.AreEqual("Function", catalog.Tests.Single(x => x.Procedure == "CheckValue").Kind);
        }

        [TestMethod]
        public void CommentsStringsAndProcedureBodiesCannotAuthorizeTests()
        {
            var catalog = Discover("'@TestModule\n" +
                "Public Sub Ordinary()\n'@TestMethod\nDebug.Print \"'@TestMethod: Public Sub Fake()\"\nEnd Sub\n" +
                "Public Sub AlsoOrdinary()\nEnd Sub\n" +
                "Rem @TestMethod\nPublic Sub RemHidden()\nEnd Sub\n" +
                "Dim Value As Long '@TestMethod\nPublic Sub InlineHidden()\nEnd Sub\n" +
                "'@TestMethod\n'Public Sub CommentedOut()\nPublic Sub Actual()\nEnd Sub");
            CollectionAssert.AreEqual(new[] { "Actual" }, catalog.Tests.Select(x => x.Procedure).ToArray());
        }

        [TestMethod]
        public void BlankLinesAndOrdinaryCommentsPreserveMetadataButStatementsTerminateIt()
        {
            var catalog = Discover("'@TestModule\n'@TestMethod\n\n' Arrange, Act, Assert\n\nPublic Sub Accepted()\nEnd Sub\n" +
                "'@TestMethod\nPrivate Value As Long\nPublic Sub Rejected()\nEnd Sub");
            Assert.AreEqual("Accepted", catalog.Tests.Single().Procedure);
            Assert.IsTrue(catalog.Diagnostics.Any(x => x.Contains("not attached")));
        }

        [TestMethod]
        public void CategoriesIgnoreAndEscapedQuotesAreCaseInsensitiveAndDeduplicated()
        {
            var test = Discover("' @testmodule\n'@testmethod\n'@TestCategory \"Arithmetic\"\n" +
                "'@TestCategory \"arithmetic\"\n'@TestCategory \"Input \"\"text\"\"\"\n'@Ignore \"Needs device\"\n" +
                "public sub Check()\nend sub").Tests.Single();
            CollectionAssert.AreEqual(new[] { "Arithmetic", "Input \"text\"" }, test.Categories);
            Assert.AreEqual("Needs device", test.IgnoreReason);
            Assert.IsNull(test.Diagnostic);
        }

        [DataTestMethod]
        [DataRow("Private Sub Check()")]
        [DataRow("Sub Check()")]
        [DataRow("Public Sub Check(value As Long)")]
        [DataRow("Public Function Check() As Long")]
        [DataRow("Public Function Check()")]
        [DataRow("Public Function Check$()")]
        [DataRow("Public Sub Check")]
        [DataRow("Public Static Sub Check()")]
        [DataRow("Public Property Get Check() As Boolean")]
        [DataRow("Public Declare Function Check Lib \"sample\" () As Boolean")]
        [DataRow("Public Sub Check() As Boolean")]
        public void UnsupportedAnnotatedSignaturesRemainVisibleAndBlocked(string signature)
        {
            var catalog = Discover("'@TestModule\n'@TestMethod\n" + signature + "\nEnd Sub");
            Assert.AreEqual(1, catalog.Tests.Count());
            StringAssert.Contains(catalog.Tests.Single().Diagnostic, "require");
        }

        [TestMethod]
        public void PhysicalLineMappingSupportsContinuationsAndColonStatements()
        {
            var catalog = Discover("'@TestModule\n'@TestMethod\nPublic Function Check( _\n) As _\nBoolean: Check = True: End Function\n" +
                "'@TestMethod\nPublic Sub Other(): End Sub");
            var test = catalog.Tests.Single(x => x.Procedure == "Check");
            Assert.AreEqual(3, test.Line);
            Assert.IsNull(test.Diagnostic);
            Assert.AreEqual(7, catalog.Tests.Single(x => x.Procedure == "Other").Line);
            Assert.IsNull(catalog.Modules.Single().Diagnostic);
        }

        [TestMethod]
        public void ConditionalTestsAndConditionalModuleOptInAreBlocked()
        {
            var catalog = Discover("#If VBA7 Then\n'@TestModule\n'@TestMethod\nPublic Sub First()\nEnd Sub\n" +
                "#Else\n'@TestMethod\nPublic Sub Second()\nEnd Sub\n#End If\n" +
                "'@TestMethod\nPublic Sub Outside()\nEnd Sub");
            StringAssert.Contains(catalog.Modules.Single().Diagnostic, "Conditional @TestModule");
            foreach (var test in catalog.Tests.Where(x => x.Procedure != "Outside"))
                StringAssert.Contains(test.Diagnostic, "Conditional compilation");
            Assert.IsNull(catalog.Tests.Single(x => x.Procedure == "Outside").Diagnostic);
        }

        [TestMethod]
        public void OptionPrivateAndNonstandardModulesAreBlocked()
        {
            var catalog = Discover("Option Private _\nModule\n'@TestModule\n'@TestMethod\nPublic Sub Check(): End Sub", 2);
            StringAssert.Contains(catalog.Modules.Single().Diagnostic, "Option Private Module");
            StringAssert.Contains(catalog.Modules.Single().Diagnostic, "standard modules");
        }

        [TestMethod]
        public void MissingModuleAndLateOptInDoNotAuthorizeTests()
        {
            var catalog = Discover("'@TestMethod\nPublic Sub Check(): End Sub\n'@TestModule\n");
            StringAssert.Contains(catalog.Modules.Single().Diagnostic, "Missing @TestModule");
            Assert.IsTrue(catalog.Diagnostics.Any(x => x.Contains("declaration section")));
        }

        [TestMethod]
        public void DuplicateAndConflictingRoleAnnotationsNeverCreateDuplicateExecution()
        {
            var catalog = Discover("'@TestModule\n'@TestModule\n'@TestMethod\n'@TestMethod\nPublic Sub One(): End Sub\n" +
                "'@TestMethod\n'@TestInitialize\nPublic Sub Two(): End Sub\n" +
                "'@ModuleCleanup\nPublic Sub Cleanup(): End Sub\n'@ModuleCleanup\nPublic Sub CleanupAgain(): End Sub");
            Assert.AreEqual(2, catalog.Tests.Count());
            Assert.IsTrue(catalog.Tests.All(x => x.Diagnostic.Contains("conflicting")));
            StringAssert.Contains(catalog.Modules.Single().Diagnostic, "Duplicate @TestModule");
            StringAssert.Contains(catalog.Modules.Single().Diagnostic, "Multiple @ModuleCleanup");
            Assert.IsNull(catalog.Modules.Single().TestInitialize);
        }

        [TestMethod]
        public void InvalidFixtureBlocksModuleAndCannotBeIgnored()
        {
            var catalog = Discover("'@TestModule\n'@TestInitialize\n'@Ignore \"Temporarily\"\n'@TestCategory \"Setup\"\n" +
                "Public Function Setup() As Boolean\nEnd Function\n'@TestMethod\nPublic Sub Check(): End Sub");
            StringAssert.Contains(catalog.Modules.Single().Diagnostic, "Fixtures require");
            StringAssert.Contains(catalog.Modules.Single().Diagnostic, "not supported on fixtures");
        }

        [DataTestMethod]
        [DataRow("'@TestCategory Arithmetic")]
        [DataRow("'@TestCategory \"\"")]
        [DataRow("'@Ignore")]
        [DataRow("'@Ignore \"unterminated")]
        [DataRow("'@Ignore \"Reason\" trailing")]
        [DataRow("'@TestMethod unexpected")]
        public void MalformedAnnotationArgumentsBlockTheTest(string annotation)
        {
            var catalog = Discover("'@TestModule\n" + (annotation.StartsWith("'@TestMethod") ? "" : "'@TestMethod\n") +
                annotation + "\nPublic Sub Check(): End Sub");
            Assert.IsFalse(string.IsNullOrEmpty(catalog.Tests.Single().Diagnostic));
        }

        [TestMethod]
        public void DuplicateProcedureNamesAreCaseInsensitiveAndBlocked()
        {
            var catalog = Discover("'@TestModule\n'@TestMethod\nPublic Sub Check(): End Sub\n'@TestMethod\nPublic Sub CHECK(): End Sub");
            Assert.AreEqual(2, catalog.Tests.Count());
            Assert.IsTrue(catalog.Tests.All(x => x.Diagnostic.Contains("Duplicate procedure identity")));
        }

        [TestMethod]
        public void AnnotatedProcedureCollidingWithOrdinaryCodeCannotExecute()
        {
            var catalog = Discover("'@TestModule\n'@TestMethod\nPublic Sub Check(): End Sub\nPublic Sub CHECK(): End Sub");
            Assert.AreEqual(1, catalog.Tests.Count());
            StringAssert.Contains(catalog.Tests.Single().Diagnostic, "Duplicate procedure identity");
        }

        [TestMethod]
        public void IdentityIsStableAcrossNameCaseAndLineEditsButDistinctAcrossProjectsAndModules()
        {
            var original = Discover("'@TestModule\n'@TestMethod\nPublic Sub Check(): End Sub").Tests.Single();
            var project = Snapshot("'@TestModule\n'@testmethod\npublic sub CHECK(): end sub");
            project.Id = "SESSION/PROJECT";
            project.Modules[0].Name = "TESTS";
            Assert.AreEqual(original.Id, VbaTestDiscovery.Discover(project).Tests.Single().Id);
            project.Id = "other/project";
            Assert.AreNotEqual(original.Id, VbaTestDiscovery.Discover(project).Tests.Single().Id);
            project.Id = "session/project";
            project.Modules[0].Name = "Other";
            Assert.AreNotEqual(original.Id, VbaTestDiscovery.Discover(project).Tests.Single().Id);
            project.Modules[0].Name = "Tests";
            project.Modules[0].Source = "\n" + project.Modules[0].Source;
            Assert.AreEqual(original.Id, VbaTestDiscovery.Discover(project).Tests.Single().Id);
            Assert.AreEqual(64, original.Id.Length);
        }

        [TestMethod]
        public void NullSourcesEmptyProjectsAndUnterminatedBodiesAreConservative()
        {
            Assert.ThrowsException<ArgumentNullException>(() => VbaTestDiscovery.Discover(null));
            Assert.AreEqual(0, VbaTestDiscovery.Discover(new VbaTestProjectSnapshot { Modules = null }).Tests.Count());
            Assert.AreEqual(0, Discover(null).Modules.Count);
            var catalog = Discover("'@TestModule\n'@TestMethod\nPublic Sub Broken()\n'@TestMethod\nPublic Sub Fake()");
            Assert.AreEqual(1, catalog.Tests.Count());
            StringAssert.Contains(catalog.Modules.Single().Diagnostic, "Unterminated");
        }

        [TestMethod]
        public void DuplicateModuleNamesCannotCreateAnAmbiguousRunnableScope()
        {
            var project = Snapshot("'@TestModule\n'@TestMethod\nPublic Sub Check(): End Sub");
            project.Modules = new[] { project.Modules[0], new VbaTestModuleSnapshot
            { Name = "TESTS", Source = project.Modules[0].Source, ComponentType = 1 } };
            Assert.IsTrue(VbaTestDiscovery.Discover(project).Modules.All(x => x.Diagnostic.Contains("Duplicate module identity")));
        }

        private static VbaTestCatalog Discover(string source, int type = 1) => VbaTestDiscovery.Discover(Snapshot(source, type));
        private static VbaTestProjectSnapshot Snapshot(string source, int type = 1) => new VbaTestProjectSnapshot
        {
            Id = "session/project", Selector = "path", Name = "Project", Revision = "revision",
            Modules = new[] { new VbaTestModuleSnapshot { Name = "Tests", Source = source, Hash = "hash", ComponentType = type } }
        };
        [TestMethod]
        public void MissingIdentitiesAndConflictingFixturesRemainConservative()
        {
            var project = Snapshot("'@TestModule\n'@ModuleInitialize\n'@ModuleCleanup\nPublic Sub Setup(): End Sub\n'@TestMethod\nPublic Sub: End Sub\n");
            project.Id = null; project.Modules[0].Name = null;
            project.Modules = new[] { null, project.Modules[0], new VbaTestModuleSnapshot { Name = null, ComponentType = 1, Source = "'@TestModule" } };
            var catalog = VbaTestDiscovery.Discover(project);
            Assert.AreEqual(2, catalog.Modules.Count);
            Assert.IsTrue(catalog.Modules.All(module => module.Diagnostic.Contains("Duplicate module identity")));
            Assert.IsTrue(catalog.Diagnostics.Any(message => message.Contains("conflicting")));
            Assert.AreEqual("<missing>", catalog.Tests.Single().Procedure);
        }

        [DataTestMethod]
        [DataRow("ModuleInitialize")]
        [DataRow("TestInitialize")]
        [DataRow("TestCleanup")]
        public void DuplicateFixturesRetainFirstDefinitionAndBlockTheModule(string role)
        {
            var catalog = Discover("'@TestModule\n'@" + role + "\nPublic Sub One(): End Sub\n'@" + role + "\nPublic Sub Two(): End Sub");
            StringAssert.Contains(catalog.Modules.Single().Diagnostic, "Multiple @" + role);
        }

        [TestMethod]
        public void MarkerArgumentsWhitespaceMetadataAndUnknownCommentsNeverGrantExecution()
        {
            var catalog = Discover("'@TestModule argument\n'@Unknown ignored\n'@TestMethod\n'@Ignore \" \"\nPublic Sub Check(): End Sub\n'@Ignore \"One\"\n'@Ignore \"Two\"\n'@TestMethod\nPublic Sub Other(): End Sub\n");
            StringAssert.Contains(catalog.Modules.Single().Diagnostic, "does not accept an argument");
            StringAssert.Contains(catalog.Tests.Single(test => test.Procedure == "Check").Diagnostic, "nonempty");
            StringAssert.Contains(catalog.Tests.Single(test => test.Procedure == "Other").Diagnostic, "Duplicate @Ignore");
            Assert.AreEqual(2, catalog.Tests.Count());
        }
        [TestMethod]
        public void MalformedDirectivesPtrSafeDeclarationsAndPropertyTerminatorsDoNotAuthorizeCalls()
        {
            var catalog = Discover("'@TestModule\n#\nEnd\nEnd Property\n'@TestMethod\nPublic Declare PtrSafe Function Native Lib \"x\" () As Boolean\n");
            Assert.AreEqual(1,catalog.Tests.Count()); StringAssert.Contains(catalog.Tests.Single().Diagnostic,"require");
        }
    }
}
