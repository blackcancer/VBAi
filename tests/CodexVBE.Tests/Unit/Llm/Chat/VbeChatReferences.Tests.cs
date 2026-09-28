namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Linq;
    using System.Text;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    [TestClass]
    [TestCategory("Unit")]
    public sealed partial class VbeChatReferenceTests
    {
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
            var module = new VbeChatReference
            {
                Project = "VBAProject",
                Module = "Module1",
                Kind = "Module"
            };
            string resolved = references.Resolve(module);
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
            var references=new VbeChatReferences(null);
            references.Execute=request=>Response.Failure("host unavailable");references.Refresh();
            Assert.AreEqual("host unavailable",references.Error);Assert.IsFalse(references.IsLoading);
            references.Step();
            references.Execute=request=>Response.Success("not an array");references.Refresh();Assert.AreEqual(0,references.Entries.Count);
            references.Execute=request=>Response.Success(new[] {new {Name="P"}});references.Refresh();
            references.Execute=request=>Response.Failure("module enumeration failed");references.Step();
            StringAssert.Contains(references.Error,"P : module enumeration failed");
            foreach(object payload in new object[] {"unexpected",new {Other=true},new {Procedures="unexpected"},new {Sha256="version",Procedures=new object[] {null,"invalid",new {Kind=0,Declaration="Public Function F()",Name="F",StartLine=1,EndLine=2},new {Kind=0,Declaration="Sub S()",Name="S",StartLine=3,EndLine=4},new {Kind=1,Declaration="Property Let",Name="Value",StartLine=5,EndLine=6}}}})
            {
                references.Execute=request=>Response.Success(new[] {new {Name=request.Command=="list_projects" ? "P":"M"}});
                references.Refresh();references.Step();
                references.Execute=request=>Response.Success(payload);references.Step();
                Assert.IsFalse(references.IsLoading);Assert.IsNull(references.Error);
                if(payload is string)Assert.AreEqual("",references.Entries[1].Sha256);
                if(references.Entries.Count==5)
                {
                    Assert.AreEqual("Function",references.Entries[2].Kind);Assert.AreEqual("Sub",references.Entries[3].Kind);
                    Assert.AreEqual("Property",references.Entries[4].Kind);Assert.AreEqual("version",references.Entries[4].Sha256);
                }
                references.Step();
            }
            references.Execute=request=>Response.Success(new[] {new {Name=request.Command=="list_projects" ? "P":"M"}});
            references.Refresh();references.Step();
            references.Execute=request=>Response.Failure("procedure enumeration failed");references.Step();
            Assert.AreEqual("#P.M : procedure enumeration failed",references.Error);
            references.Execute=request=>Response.Success(new[] {new {MissingName=true}});references.Refresh();Assert.AreEqual("#",references.Entries[0].Token);
        }

        [TestMethod]
        public void NavigationMatrixUsesFreshHashForModuleAndProcedureAndReturnsReadFailures()
        {
            var references=new VbeChatReferences(null);
            var item=new VbeChatReference {Project="P",Module="M",Kind="Module"};
            references.Execute=request=>Response.Failure("module disappeared");
            Assert.AreEqual("module disappeared",references.Navigate(item).Error);
            Assert.ThrowsException<InvalidOperationException>(()=>references.Resolve(item));
            string command=null;
            references.Execute=request=>
            {
                if(request.Command=="read_module")return Response.Success(new {Code="line1\nline2",Sha256="fresh"});
                command=request.Command;Assert.AreEqual("fresh",request.ExpectedSha256);Assert.AreEqual(1,request.StartLine);
                return Response.Success(new {Selected=true});
            };
            Assert.IsTrue(references.Navigate(item).Ok);Assert.AreEqual("select_code",command);
            item.Name="F";item.ProcKind=3;
            Assert.IsTrue(references.Navigate(item).Ok);Assert.AreEqual("select_procedure",command);
            item.Sha256="FRESH";item.StartLine=-1;item.EndLine=99;
            StringAssert.Contains(references.Resolve(item),"line1\nline2");
            references.Execute=request=>Response.Success(new {MissingCode=true});item.Name=null;
            StringAssert.Contains(references.Resolve(item),"SHA-256");
        }

        [TestMethod]
        public void ReferenceMatchingMatrixLimitsSortsAndDisplaysAllTokenKinds()
        {
            using(var localization=new Infrastructure.LocalizationScope())
            {
                var references=new VbeChatReferences(null);
                foreach(int kind in new[] {1,2,3})
                {
                    var item=new VbeChatReference {Project="P",Module="M",Name="Value",Kind="Property",ProcKind=kind};
                    Assert.AreEqual(new[] {"","Let","Set","Get"}[kind],item.Token.Split(':').Last());
                    Assert.AreEqual(item.Token,item.DisplayToken);Assert.AreEqual(item.Display,item.ToString());
                    StringAssert.Contains(item.Token,item.Name);
                    references.Entries.Add(item);
                }
                var project=new VbeChatReference {Project="P",Kind="Projet"};
                Assert.AreEqual(UiText.Get("Project"),project.DisplayKind);
                StringAssert.Contains(project.Display,project.DisplayKind);
                references.Entries.Add(project);
                references.Entries.Add(new VbeChatReference {Project="Other",Module="M",Name="PInside",Kind="Sub"});
                Assert.AreEqual(5,references.MatchPrefix("P",'\0').Count());
                Assert.AreEqual(3,references.MatchPrefix("Value",'@').Count());
                Assert.AreEqual(3,references.MatchPrefix("vAlUe",'@').Count());
                Assert.AreEqual(3,references.Match("VALUE").Count());
                Assert.AreEqual(0,references.MatchPrefix("no-match",'@').Count());
                Assert.AreEqual(1,references.MatchPrefix("P",'#').Count());
                Assert.AreEqual("#P",references.Match("P").First().Token);
                references.Entries.Clear();
                for(int i=0;i<45;i++)references.Entries.Add(new VbeChatReference {Project="P"+i.ToString("D2"),Kind="Projet"});
                Assert.AreEqual(40,references.Match("").Count());Assert.AreEqual("#P00",references.Match("").First().Token);
            }
        }

    }
}
