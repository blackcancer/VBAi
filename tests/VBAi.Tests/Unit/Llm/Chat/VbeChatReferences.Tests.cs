namespace VBAi.Tests.Unit
{
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using System;
    using System.Linq;
    using System.Text;
    using VBAi;

    [TestClass]
    [TestCategory("Unit")]
    public sealed partial class VbeChatReferenceTests
    {
        /// <summary>Every scalar field survives capture; mutations of discovery objects cannot change the snapshot.</summary>
        [TestMethod]
        public void ReferenceSnapshotPreservesAllCapturedScalarFieldsWithoutSharingMutableIdentity()
        {
            var original = new VbeChatReference { Project = "P", Module = "M", Name = "Value", Kind = "Property", ProcKind = 3,
                DeclarationScope = "Member", DeclarationColumn = 9, StartLine = 17, EndLine = 23, Sha256 = "old" };
            var snapshot = original.CaptureSnapshot();
            var json = new System.Web.Script.Serialization.JavaScriptSerializer();
            Assert.AreNotSame(original, snapshot); Assert.AreEqual(json.Serialize(original), json.Serialize(snapshot));
            string token = snapshot.Token;
            original.Project = "Changed"; original.Module = "Other"; original.Name = "Different"; original.Kind = "Sub";
            original.ProcKind = 0; original.StartLine = 99; original.EndLine = 100; original.DeclarationColumn = 1; original.Sha256 = "new";
            Assert.AreEqual(token, snapshot.Token); Assert.AreEqual("old", snapshot.Sha256);
            Assert.AreEqual(17, snapshot.StartLine); Assert.AreEqual(23, snapshot.EndLine); Assert.AreEqual(9, snapshot.DeclarationColumn);
        }

        [TestMethod]
        public void LiveContextNamesHostAndRealAnsiCodePageWithoutAssumingExcel()
        {
            dynamic snapshot = LlmVbeContext.LiveSnapshot(Session());
            Assert.IsTrue((bool)snapshot.VbeConnected);
            Assert.AreEqual(Encoding.Default.CodePage, (int)snapshot.HostAnsiCodePage);
            Assert.IsNotNull(snapshot.Projects);
            StringAssert.Contains(LlmVbeContext.DeveloperInstructions, "never assume Excel");
            StringAssert.Contains(LlmVbeContext.EncodingInstructions, Encoding.Default.CodePage.ToString());
        }

        [TestMethod]
        public void ReferenceDiscoveryLoadsProjectsThenModulesAndReportsProcedureProbeError()
        {
            var references = new VbeChatReferences(Session());
            int changed = 0;
            references.Changed += () => changed++;
            references.Refresh();
            Assert.AreEqual(1, references.Entries.Count);
            Assert.AreEqual("#VBAProject", references.Entries[0].Token);
            Assert.IsTrue(references.IsLoading);
            references.Step();
            Assert.AreEqual(2, references.Entries.Count);
            Assert.AreEqual("#VBAProject.Module1", references.Entries[1].Token);
            Assert.IsTrue(references.IsLoading);
            references.Step();
            Assert.IsFalse(references.IsLoading);
            Assert.IsFalse(string.IsNullOrWhiteSpace(references.Error));
            Assert.AreEqual(3, changed);
        }

        [TestMethod]
        public void TokensFilterProjectModuleAndPropertyProcedurePrefixes()
        {
            var references = new VbeChatReferences(Session());
            references.Refresh();
            references.Step();
            references.Entries.Add(new VbeChatReference { Project = "VBAProject", Module = "Module1", Name = "Title", Kind = "Property", ProcKind = 3 });
            Assert.AreEqual("@VBAProject.Module1.Title:Get", references.Entries.Last().Token);
            Assert.AreEqual(2, references.MatchPrefix("#VBAProject", '#').Count());
            Assert.AreEqual(1, references.MatchPrefix("@Title", '@').Count());
            Assert.AreEqual(0, references.MatchPrefix("@Title", '#').Count());
            Assert.AreEqual(1, references.Match("Title").Count());
        }

        [TestMethod]
        public void ResolveUsesFreshModuleSourceAndRefusesStaleProcedureSelection()
        {
            var references = new VbeChatReferences(Session());
            var project = new VbeChatReference
            {
                Project = "VBAProject",
                Kind = "Projet"
            };
            StringAssert.Contains(references.Resolve(project), "Module1");
            references.Refresh();
            while (references.IsLoading) references.Step();
            var module = references.CaptureSelection(references.Entries.Single(item => item.Module == "Module1" && item.Name == null));
            string selectedSha = module.Sha256;
            string resolved = references.Resolve(module);
            Assert.AreEqual(selectedSha, module.Sha256, "Resolution must preserve the selected revision.");
            StringAssert.Contains(resolved, "Sub Example()");
            Assert.IsFalse(string.IsNullOrWhiteSpace(module.Sha256));
            var procedure = new VbeChatReference
            {
                Project = "VBAProject",
                Module = "Module1",
                Name = "Example",
                Kind = "Sub",
                StartLine = 1,
                EndLine = 2,
                Sha256 = "stale"
            };
            Assert.ThrowsException<InvalidOperationException>(() => references.Resolve(procedure));
            procedure.Sha256 = module.Sha256;
            StringAssert.Contains(references.Resolve(procedure), "End Sub");
            procedure.StartLine = 99;
            Assert.ThrowsException<InvalidOperationException>(() => references.Resolve(procedure));
        }

        [TestMethod]
        public void DeclarationReferencesKeepScopesAndRefuseStaleSourceBeforeNavigation()
        {
            var references = new VbeChatReferences(null);
            Request selected = null;
            references.Execute = request =>
            {
                if (request.Command == "list_projects") return Response.Success(new[] { new { Name = "P" } });
                if (request.Command == "list_modules") return Response.Success(new[] { new { Name = "M" } });
                if (request.Command == "list_procedures") return Response.Success(new { Sha256 = "fresh", Procedures = new object[0] });
                if (request.Command == "read_module") return Response.Success(new { Sha256 = "fresh", Code = "Sub Run()\nDim local As Long\nEnd Sub" });
                selected = request; return Response.Success(new { Selected = true });
            };
            references.Refresh(); while (references.IsLoading) references.Step();
            var declaration = references.Entries.Single(x => x.Name == "local");
            Assert.AreEqual("@P.M.Run.local:L2", declaration.Token);
            Assert.AreEqual(1, references.MatchPrefix("local", '@').Count());
            Assert.IsTrue(references.Navigate(declaration).Ok);
            Assert.AreEqual("select_code", selected.Command); Assert.AreEqual(2, selected.StartLine);
            Assert.AreEqual("local", selected.Expression);
            StringAssert.Contains(references.Resolve(declaration), "Dim local As Long");
            declaration.Sha256 = "stale"; selected = null;
            Assert.IsFalse(references.Navigate(declaration).Ok); Assert.IsNull(selected);
        }

        [TestMethod]
        public void ProjectReferenceCannotNavigateDirectlyToCode()
        {
            var references = new VbeChatReferences(Session());
            var result = references.Navigate(new VbeChatReference { Project = "VBAProject", Kind = "Projet" });
            Assert.IsFalse(result.Ok);
            StringAssert.Contains(result.Error, "module");
        }
        [TestMethod]
        public void ReferenceDiscoveryMatrixHandlesProcedureShapesErrorsAndOptionalObservers()
        {
            var references = new VbeChatReferences(null);
            references.Execute = request => Response.Failure("host unavailable"); references.Refresh();
            Assert.AreEqual("host unavailable", references.Error); Assert.IsFalse(references.IsLoading);
            references.Step();
            references.Execute = request => Response.Success("not an array"); references.Refresh(); Assert.AreEqual(0, references.Entries.Count);
            references.Execute = request => Response.Success(new[] { new { Name = "P" } }); references.Refresh();
            references.Execute = request => Response.Failure("module enumeration failed"); references.Step();
            StringAssert.Contains(references.Error, "P : module enumeration failed");
            foreach (object payload in new object[] { "unexpected", new { Other = true }, new { Procedures = "unexpected" }, new { Sha256 = "version", Procedures = new object[] { null, "invalid", new { Kind = 0, Declaration = "Public Function F()", Name = "F", StartLine = 1, EndLine = 2 }, new { Kind = 0, Declaration = "Sub S()", Name = "S", StartLine = 3, EndLine = 4 }, new { Kind = 1, Declaration = "Property Let", Name = "Value", StartLine = 5, EndLine = 6 } } } })
            {
                references.Execute = request => Response.Success(new[] { new { Name = request.Command == "list_projects" ? "P" : "M" } });
                references.Refresh(); references.Step();
                references.Execute = request => Response.Success(payload); references.Step();
                Assert.IsFalse(references.IsLoading); Assert.IsNull(references.Error);
                if (payload is string) Assert.AreEqual("", references.Entries[1].Sha256);
                if (references.Entries.Count == 5)
                {
                    Assert.AreEqual("Function", references.Entries[2].Kind); Assert.AreEqual("Sub", references.Entries[3].Kind);
                    Assert.AreEqual("Property", references.Entries[4].Kind); Assert.AreEqual("version", references.Entries[4].Sha256);
                }
                references.Step();
            }
            references.Execute = request => Response.Success(new[] { new { Name = request.Command == "list_projects" ? "P" : "M" } });
            references.Refresh(); references.Step();
            references.Execute = request => Response.Failure("procedure enumeration failed"); references.Step();
            Assert.AreEqual("#P.M : procedure enumeration failed", references.Error);
            references.Execute = request => Response.Success(new[] { new { MissingName = true } }); references.Refresh(); Assert.AreEqual("#", references.Entries[0].Token);
        }

        [TestMethod]
        public void NavigationMatrixUsesFreshHashForModuleAndProcedureAndReturnsReadFailures()
        {
            var references = new VbeChatReferences(null);
            var item = new VbeChatReference { Project = "P", Module = "M", Kind = "Module" };
            references.Execute = request => Response.Failure("module disappeared");
            Assert.AreEqual("module disappeared", references.Navigate(item).Error);
            Assert.ThrowsException<InvalidOperationException>(() => references.Resolve(item));
            string command = null;
            references.Execute = request =>
            {
                if (request.Command == "read_module") return Response.Success(new { Code = "line1\nline2", Sha256 = "fresh" });
                command = request.Command; Assert.AreEqual("fresh", request.ExpectedSha256); Assert.AreEqual(1, request.StartLine);
                return Response.Success(new { Selected = true });
            };
            Assert.IsTrue(references.Navigate(item).Ok); Assert.AreEqual("select_code", command);
            item.Name = "F"; item.ProcKind = 3;
            Assert.IsTrue(references.Navigate(item).Ok); Assert.AreEqual("select_procedure", command);
            item.Sha256 = "FRESH"; item.StartLine = -1; item.EndLine = 99;
            StringAssert.Contains(references.Resolve(item), "line1\nline2");
            references.Execute = request => Response.Success(new { MissingCode = true }); item.Name = null;
            Assert.ThrowsException<InvalidOperationException>(() => references.Resolve(item));
            Assert.AreEqual("FRESH", item.Sha256, "A missing live revision must not overwrite the selected one.");
            references.Execute = request => Response.Success(new { Code = "", Sha256 = "fresh" });
            StringAssert.Contains(references.Resolve(item), "SHA-256 fresh");
            Assert.AreEqual("FRESH", item.Sha256);
        }

        [TestMethod]
        public void ReferenceMatchingMatrixLimitsSortsAndDisplaysAllTokenKinds()
        {
            using (var localization = new Infrastructure.LocalizationScope())
            {
                var references = new VbeChatReferences(null);
                foreach (int kind in new[] { 1, 2, 3 })
                {
                    var item = new VbeChatReference { Project = "P", Module = "M", Name = "Value", Kind = "Property", ProcKind = kind };
                    Assert.AreEqual(new[] { "", "Let", "Set", "Get" }[kind], item.Token.Split(':').Last());
                    Assert.AreEqual(item.Token, item.DisplayToken); Assert.AreEqual(item.Display, item.ToString());
                    StringAssert.Contains(item.Token, item.Name);
                    references.Entries.Add(item);
                }
                var project = new VbeChatReference { Project = "P", Kind = "Projet" };
                Assert.AreEqual(UiText.Get("Project"), project.DisplayKind);
                StringAssert.Contains(project.Display, project.DisplayKind);
                references.Entries.Add(project);
                references.Entries.Add(new VbeChatReference { Project = "Other", Module = "M", Name = "PInside", Kind = "Sub" });
                Assert.AreEqual(5, references.MatchPrefix("P", '\0').Count());
                Assert.AreEqual(3, references.MatchPrefix("Value", '@').Count());
                Assert.AreEqual(3, references.MatchPrefix("vAlUe", '@').Count());
                Assert.AreEqual(3, references.Match("VALUE").Count());
                Assert.AreEqual(0, references.MatchPrefix("no-match", '@').Count());
                Assert.AreEqual(1, references.MatchPrefix("P", '#').Count());
                Assert.AreEqual("#P", references.Match("P").First().Token);
                references.Entries.Clear();
                for (int i = 0; i < 45; i++) references.Entries.Add(new VbeChatReference { Project = "P" + i.ToString("D2"), Kind = "Projet" });
                Assert.AreEqual(40, references.Match("").Count()); Assert.AreEqual("#P00", references.Match("").First().Token);
            }
        }


        [TestMethod]
        public void DeclarationDiscoveryMarksModuleScopeAndReportsSourceFailures()
        {
            foreach (bool fail in new[] { false, true })
            {
                var references = new VbeChatReferences(null);
                references.Execute = request =>
                {
                    switch (request.Command)
                    {
                        case "list_projects": return Response.Success(new[] { new { Name = "P" } });
                        case "list_modules": return Response.Success(new[] { new { Name = "M" } });
                        case "list_procedures": return Response.Success(new { Sha256 = "version", Procedures = new object[0] });
                        case "read_module": return fail ? Response.Failure("source unavailable") : Response.Success(new { Sha256 = "version", Code = "Dim moduleValue As Long" });
                        default: throw new InvalidOperationException(request.Command);
                    }
                };
                references.Refresh(); while (references.IsLoading) references.Step();
                if (fail) StringAssert.Contains(references.Error, "source unavailable");
                else { var declaration = references.Entries.Single(x => x.Name == "moduleValue"); Assert.IsNull(declaration.DeclarationScope); Assert.AreEqual("version", declaration.Sha256); }
            }
        }
    }
}

namespace VBAi.Tests.Unit
{
    using System;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>Requires the selected source revision for both modules and procedures.</summary>
    [TestClass, TestCategory("Unit")]
    public sealed class VbeChatReferenceRevisionTests
    {
        /// <summary>Pins only an unpinned whole module and detaches selection metadata from ongoing discovery.</summary>
        [DataTestMethod]
        [DataRow("project", null, null, false, true, 0)]
        [DataRow("module", "old", "new", false, true, 0)]
        [DataRow("module", null, "fresh", false, true, 1)]
        [DataRow("module", "", "fresh", false, true, 1)]
        [DataRow("module", " ", "fresh", false, true, 1)]
        [DataRow("module", null, null, false, false, 1)]
        [DataRow("module", null, " ", false, false, 1)]
        [DataRow("module", null, "fresh", true, false, 1)]
        [DataRow("procedure", null, "fresh", false, false, 0)]
        [DataRow("procedure", "old", "new", false, true, 0)]
        public void SelectionPinsLoadingModulesWithoutReplacingKnownRevisions(
            string kind, string selectedRevision, string liveRevision, bool readFailure, bool accepted, int expectedReads)
        {
            int reads = 0;
            var original = new VbeChatReference
            {
                Project = "P", Module = kind == "project" ? null : "M",
                Name = kind == "procedure" ? "Value" : null,
                Sha256 = selectedRevision, Kind = kind, StartLine = 2, EndLine = 3
            };
            var references = new VbeChatReferences(null);
            references.Execute = request =>
            {
                reads++; Assert.AreEqual("read_module", request.Command);
                Assert.AreEqual("P", request.Project); Assert.AreEqual("M", request.Module);
                original.Project = "discovery update"; original.Sha256 = "discovery update";
                return readFailure ? Response.Failure("read refused") : Response.Success(new { Code = "private source", Sha256 = liveRevision });
            };
            if (accepted)
            {
                var snapshot = references.CaptureSelection(original);
                Assert.AreNotSame(original, snapshot);
                Assert.AreEqual("P", snapshot.Project);
                Assert.AreEqual(kind == "project" ? null : "M", snapshot.Module);
                Assert.AreEqual(expectedReads == 0 ? selectedRevision : liveRevision, snapshot.Sha256);
                Assert.AreEqual(2, snapshot.StartLine); Assert.AreEqual(3, snapshot.EndLine);
                original.Project = "changed later"; original.Sha256 = "changed later";
                Assert.AreEqual("P", snapshot.Project);
                Assert.AreEqual(expectedReads == 0 ? selectedRevision : liveRevision, snapshot.Sha256);
            }
            else
            {
                var error = Assert.ThrowsException<InvalidOperationException>(() => references.CaptureSelection(original));
                Assert.IsFalse(error.Message.Contains("private source"));
            }
            Assert.AreEqual(expectedReads, reads);
            Assert.ThrowsException<ArgumentNullException>(() => references.CaptureSelection(null));
        }

        [DataTestMethod]
        [DataRow(false, "hash", "hash", true)]
        [DataRow(true, "hash", "HASH", true)]
        [DataRow(false, "old", "new", false)]
        [DataRow(true, "old", "new", false)]
        [DataRow(false, null, "hash", false)]
        [DataRow(false, "", "hash", false)]
        [DataRow(true, null, "hash", false)]
        [DataRow(true, "", "hash", false)]
        [DataRow(false, "hash", null, false)]
        [DataRow(false, "hash", "", false)]
        [DataRow(true, "hash", null, false)]
        [DataRow(true, "hash", "", false)]
        [DataRow(false, null, null, false)]
        [DataRow(true, null, null, false)]
        [DataRow(false, "hash", "HASH", true)]
        [DataRow(true, "hash", "hash", true)]
        public void ResolutionPreservesSelectionAndRejectsStaleOrMissingRevisions(
            bool procedure, string selectedHash, string liveHash, bool accepted)
        {
            int reads = 0;
            var references = new VbeChatReferences(null);
            references.Execute = request =>
            {
                reads++;
                Assert.AreEqual("read_module", request.Command);
                Assert.AreEqual("P", request.Project);
                Assert.AreEqual("M", request.Module);
                return Response.Success(new { Code = "first\r\nsecond\r\nthird", Sha256 = liveHash });
            };
            var item = new VbeChatReference
            {
                Project = "P", Module = "M", Name = procedure ? "Value" : null,
                Kind = procedure ? "Sub" : "Module", StartLine = 2, EndLine = 2,
                Sha256 = selectedHash
            };
            if (accepted)
            {
                string resolved = references.Resolve(item);
                StringAssert.Contains(resolved, "second");
                if (procedure)
                {
                    Assert.IsFalse(resolved.Contains("first"));
                    Assert.IsFalse(resolved.Contains("third"));
                }
                else StringAssert.Contains(resolved, "first\r\nsecond\r\nthird");
            }
            else
            {
                var error = Assert.ThrowsException<InvalidOperationException>(() => references.Resolve(item));
                StringAssert.Contains(error.Message, item.Token);
                Assert.IsFalse(error.Message.Contains("first\r\nsecond\r\nthird"));
            }
            Assert.AreEqual(1, reads);
            Assert.AreEqual(selectedHash, item.Sha256);
            Assert.AreEqual(2, item.StartLine);
            Assert.AreEqual(2, item.EndLine);
        }
    }
}
