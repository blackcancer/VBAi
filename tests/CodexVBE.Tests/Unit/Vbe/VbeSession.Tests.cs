namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    [TestClass]
    [TestCategory("Unit")]
    public sealed partial class VbeSessionContractTests
    {
        [TestMethod]
        public void DispatcherRejectsMissingAndUnknownCommandsWithoutTouchingTheHost()
        {
            var session = new VbeSession(new FakeVbe());
            Assert.AreEqual("A command is required.", session.Execute(null).Error);
            Assert.AreEqual("A command is required.", session.Execute(new Request { Command = " " }).Error);
            Assert.AreEqual("Unknown command: no_such_command", session.Execute(new Request { Command = "no_such_command" }).Error);
            Assert.IsTrue(session.Execute(new Request { Command = "status" }).Ok);
        }

        [TestMethod]
        public void ProjectAndModuleDiscoveryRetainNamesTypesAndLineCounts()
        {
            var project = new FakeProject
            {
                Name = "ClasseurÉté",
                FileName = @"C:\Temp\été.xlsm",
                Mode = 2
            };
            project.VBComponents.Add(new FakeComponent { Name = "Calcul", Type = 1, CodeModule = new FakeModule { CountOfLines = 7 } });
            project.VBComponents.Add(new FakeComponent { Name = "Métier", Type = 2, CodeModule = new FakeModule { CountOfLines = 3 } });
            var host = new FakeVbe();
            host.VBProjects.Add(project);
            var session = new VbeSession(host);
            var projects = (IEnumerable<object>)session.Execute(new Request { Command = "list_projects" }).Data;
            dynamic discoveredProject = projects.Single();
            Assert.AreEqual("ClasseurÉté", (string)discoveredProject.Name);
            Assert.AreEqual(project.FileName, (string)discoveredProject.FileName);
            Assert.AreEqual(2, (int)discoveredProject.Mode);
            var response = session.Execute(new Request { Command = "list_modules", Project = "classeurété" });
            Assert.IsTrue(response.Ok);
            var modules = ((IEnumerable<object>)response.Data).Cast<dynamic>().ToArray();
            Assert.AreEqual(2, modules.Length);
            Assert.AreEqual("Calcul", (string)modules[0].Name);
            Assert.AreEqual(1, (int)modules[0].Type);
            Assert.AreEqual(7, (int)modules[0].Lines);
            Assert.AreEqual("Métier", (string)modules[1].Name);
            Assert.AreEqual(2, (int)modules[1].Type);
            Assert.AreEqual(3, (int)modules[1].Lines);
        }

        [TestMethod]
        public void ReadAndEditModuleRequireTheVersionAndDesignMode()
        {
            var project = new FakeProject
            {
                Name = "Projet",
                Mode = 2
            };
            var module = new FakeModule("Option Explicit\r\nSub Essai()\r\nEnd Sub");
            project.VBComponents.Add(new FakeComponent { Name = "Module1", Type = 1, CodeModule = module });
            var host = new FakeVbe();
            host.VBProjects.Add(project);
            var session = new VbeSession(host);
            dynamic read = session.Execute(new Request { Command = "read_module", Project = "Projet", Module = "Module1" }).Data;
            Assert.AreEqual(module.Code, (string)read.Code);
            Assert.IsFalse(session.Execute(new Request { Command = "replace_lines", Project = "Projet", Module = "Module1" }).Ok);
            Assert.IsFalse(session.Execute(new Request { Command = "replace_lines", Project = "Projet", Module = "Module1", ExpectedSha256 = "outdated", StartLine = 1, Count = 1, Text = "Option Private Module" }).Ok);
            project.Mode = 1;
            Assert.IsFalse(session.Execute(new Request { Command = "replace_lines", Project = "Projet", Module = "Module1", ExpectedSha256 = (string)read.Sha256, StartLine = 1, Count = 1, Text = "Option Private Module" }).Ok);
            project.Mode = 2;
            var edit = session.Execute(new Request { Command = "replace_lines", Project = "Projet", Module = "Module1", ExpectedSha256 = (string)read.Sha256, StartLine = 1, Count = 1, Text = "Option Private Module" });
            Assert.IsTrue(edit.Ok);
            StringAssert.StartsWith(module.Code, "Option Private Module");
            Assert.IsFalse(session.Execute(new Request { Command = "replace_lines", Project = "Projet", Module = "Module1", ExpectedSha256 = (string)read.Sha256, StartLine = 1, Count = 1, Text = "Option Explicit" }).Ok);
        }
    }
}

namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Security.Cryptography;
    using System.Text;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    [TestClass]
    [TestCategory("Unit")]
    public sealed partial class VbeSessionTests
    {
        [TestMethod]
        public void DispatchReportsMissingUnknownAndStatusCommands()
        {
            var f = Create();
            foreach (var request in new[]
            {
                null,
                new Request(),
                new Request
                {
                    Command = " "
                }
            }

            )
            {
                var response = f.Session.Execute(request);
                Assert.IsFalse(response.Ok);
                StringAssert.Contains(response.Error, "command is required");
            }

            var unknown = f.Session.Execute(new Request { Command = "does_not_exist" });
            Assert.IsFalse(unknown.Ok);
            Assert.AreEqual("Unknown command: does_not_exist", unknown.Error);
            var status = f.Session.Execute(new Request { Command = "status" });
            Assert.IsTrue(status.Ok);
            Assert.IsTrue((bool)((dynamic)status.Data).Connected);
        }

        [TestMethod]
        public void ListProjectsRetainsUnsavedProjectWhenFileNameIsUnavailable()
        {
            var f = Create();
            f.Vbe.VBProjects.Add(new FakeProject { Name = "Unsaved", ThrowFileName = true, Mode = 1 });
            var response = f.Session.Execute(new Request { Command = "list_projects" });
            Assert.IsTrue(response.Ok);
            var projects = ((IEnumerable)response.Data).Cast<object>().ToArray();
            Assert.AreEqual(2, projects.Length);
            Assert.AreEqual(@"C:\Temp\Host.xlsm", (string)((dynamic)projects[0]).FileName);
            Assert.AreEqual("Unsaved", (string)((dynamic)projects[1]).Name);
            Assert.IsNull((string)((dynamic)projects[1]).FileName);
            Assert.AreEqual(1, (int)((dynamic)projects[1]).Mode);
        }

        [TestMethod]
        public void ListModulesAndReadModuleDispatchToSelectedProject()
        {
            var f = Create();
            f.Project.VBComponents.Items.Add(new FakeComponent { Name = "Class1", Type = 2, CodeModule = new FakeModule("") });
            var listed = f.Session.Execute(new Request { Command = "list_modules", Project = "vbaproject" });
            Assert.IsTrue(listed.Ok);
            var modules = ((IEnumerable)listed.Data).Cast<object>().ToArray();
            Assert.AreEqual(2, modules.Length);
            Assert.AreEqual(2, (int)((dynamic)modules[0]).Lines);
            Assert.AreEqual(0, (int)((dynamic)modules[1]).Lines);
            var read = f.Session.Execute(new Request { Command = "read_module", Project = f.Project.Name, Module = "module1" });
            Assert.IsTrue(read.Ok);
            Assert.AreEqual("Alpha\r\nBeta", (string)((dynamic)read.Data).Code);
            Assert.AreEqual(Sha("Alpha\r\nBeta"), (string)((dynamic)read.Data).Sha256);
            var empty = f.Session.Execute(new Request { Command = "read_module", Project = f.Project.Name, Module = "Class1" });
            Assert.AreEqual("", (string)((dynamic)empty.Data).Code);
            Assert.AreEqual(Sha(""), (string)((dynamic)empty.Data).Sha256);
        }

        [TestMethod]
        public void ReadModuleRejectsMissingOrUnknownModule()
        {
            var f = Create();
            Assert.ThrowsException<ArgumentException>(() => f.Session.Execute(new Request { Command = "read_module", Project = f.Project.Name }));
            Assert.ThrowsException<InvalidOperationException>(() => f.Session.Execute(new Request { Command = "read_module", Project = f.Project.Name, Module = "Missing" }));
        }

        [TestMethod]
        public void ReplaceLinesRejectsMissingHashRangeAndChangedCodeWithoutMutation()
        {
            var f = Create();
            var request = new Request
            {
                Command = "replace_lines",
                Project = f.Project.Name,
                Module = "Module1",
                StartLine = 1,
                Count = 1,
                Text = "Changed"
            };
            Assert.IsFalse(f.Session.Execute(request).Ok);
            request.ExpectedSha256 = Sha("stale");
            StringAssert.Contains(f.Session.Execute(request).Error, "module changed");
            request.ExpectedSha256 = Sha("Alpha\r\nBeta");
            request.StartLine = 4;
            StringAssert.Contains(f.Session.Execute(request).Error, "outside the module");
            request.StartLine = 1;
            request.Count = -1;
            StringAssert.Contains(f.Session.Execute(request).Error, "Invalid line range");
            Assert.AreEqual("Alpha\r\nBeta", f.Project.VBComponents.Items[0].CodeModule.Code);
        }

        [TestMethod]
        public void ReplaceLinesRequiresDesignMode()
        {
            var f = Create(mode: 1);
            var response = f.Session.Execute(new Request { Command = "replace_lines", Project = f.Project.Name, Module = "Module1", StartLine = 1, Count = 1, Text = "Changed", ExpectedSha256 = Sha("Alpha\r\nBeta") });
            Assert.IsFalse(response.Ok);
            StringAssert.Contains(response.Error, "design mode");
            Assert.AreEqual("Alpha\r\nBeta", f.Project.VBComponents.Items[0].CodeModule.Code);
        }

        [TestMethod]
        public void ReplaceLinesDeletesAndInsertsThenReportsCurrentHash()
        {
            var f = Create();
            var response = f.Session.Execute(new Request { Command = "replace_lines", Project = f.Project.Name, Module = "Module1", StartLine = 2, Count = 1, Text = "Gamma", ExpectedSha256 = Sha("Alpha\r\nBeta") });
            Assert.IsTrue(response.Ok);
            Assert.AreEqual("Alpha\r\nGamma", f.Project.VBComponents.Items[0].CodeModule.Code);
            Assert.AreEqual(Sha("Alpha\r\nGamma"), (string)((dynamic)response.Data).Sha256);
            Assert.AreEqual(2, (int)((dynamic)response.Data).Lines);
        }

        [TestMethod]
        public void ReplaceLinesRestoresDeletedSourceAfterInsertFailure()
        {
            var f = Create();
            var module = f.Project.VBComponents.Items[0].CodeModule;
            module.InsertFailuresRemaining = 1;
            var response = f.Session.Execute(new Request { Command = "replace_lines", Project = f.Project.Name, Module = "Module1", StartLine = 1, Count = 2, Text = "Changed", ExpectedSha256 = Sha(module.Code) });
            Assert.IsFalse(response.Ok);
            StringAssert.Contains(response.Error, "original source restored");
            Assert.AreEqual("Alpha\r\nBeta", module.Code);
        }

        [TestMethod]
        public void ReplaceLinesReportsFailedRollbackWithoutClaimingSuccess()
        {
            var f = Create();
            var module = f.Project.VBComponents.Items[0].CodeModule;
            module.InsertFailuresRemaining = 2;
            var response = f.Session.Execute(new Request { Command = "replace_lines", Project = f.Project.Name, Module = "Module1", StartLine = 1, Count = 2, Text = "Changed", ExpectedSha256 = Sha(module.Code) });
            Assert.IsFalse(response.Ok);
            StringAssert.Contains(response.Error, "Rollback failed");
            Assert.AreEqual("", module.Code);
        }

        [TestMethod]
        public void CreateModuleAndClassValidateModeAndDuplicateIdentity()
        {
            var f = Create();
            var request = new Request
            {
                Command = "create_module",
                Project = f.Project.Name,
                Module = "Module1",
                ExpectedMode = 2
            };
            Assert.ThrowsException<InvalidOperationException>(() => f.Session.Execute(request));
            request.Module = "2Bad";
            Assert.ThrowsException<ArgumentException>(() => f.Session.Execute(request));
            request.Module = "NewModule";
            request.ExpectedMode = 1;
            Assert.ThrowsException<ArgumentException>(() => f.Session.Execute(request));
            request.ExpectedMode = 2;
            var created = f.Session.Execute(request);
            Assert.IsTrue(created.Ok);
            Assert.AreEqual(1, (int)((dynamic)created.Data).Type);
            Assert.AreEqual(Sha(""), (string)((dynamic)created.Data).Sha256);
            request.Command = "create_class";
            request.Module = "NewClass";
            var createdClass = f.Session.Execute(request);
            Assert.AreEqual(2, (int)((dynamic)createdClass.Data).Type);
            Assert.AreEqual(3, f.Project.VBComponents.Items.Count);
        }

        [TestMethod]
        public void CreateComponentRollsBackOnlyNewItemWhenRenameFails()
        {
            var f = Create();
            f.Project.VBComponents.RejectedName = "Denied";
            var request = new Request
            {
                Command = "create_module",
                Project = f.Project.Name,
                Module = "Denied",
                ExpectedMode = 2
            };
            Assert.ThrowsException<InvalidOperationException>(() => f.Session.Execute(request));
            Assert.AreEqual(1, f.Project.VBComponents.Items.Count);
            Assert.AreEqual("Module1", f.Project.VBComponents.Items[0].Name);
            Assert.AreEqual(1, f.Project.VBComponents.RemoveCount);
        }

        [TestMethod]
        public void ReferenceSnapshotKeepsBrokenIdentityAndChangesVersionWithInventory()
        {
            var f = Create();
            var broken = new FakeReference
            {
                GUID = "{AAAAAAAA-AAAA-AAAA-AAAA-AAAAAAAAAAAA}",
                Major = 1,
                Minor = 0,
                IsBroken = true,
                ThrowMetadata = true
            };
            f.Project.References.Items.Add(broken);
            var before = f.Session.Execute(new Request { Command = "list_references", Project = f.Project.Name });
            Assert.IsTrue(before.Ok);
            dynamic snapshot = before.Data;
            var entries = ((IEnumerable)snapshot.References).Cast<object>().ToArray();
            Assert.AreEqual(1, entries.Length);
            Assert.IsTrue((bool)entries[0].GetType().GetProperty("IsBroken").GetValue(entries[0]));
            Assert.IsNull(entries[0].GetType().GetProperty("Name").GetValue(entries[0]));
            string version = (string)snapshot.Version;
            f.Project.References.Items.Add(new FakeReference { GUID = "{BBBBBBBB-BBBB-BBBB-BBBB-BBBBBBBBBBBB}", Name = "Library", FullPath = @"C:\Temp\Library.dll", Major = 2, Minor = 1 });
            var after = f.Session.Execute(new Request { Command = "list_references", Project = f.Project.Name });
            Assert.AreNotEqual(version, (string)((dynamic)after.Data).Version);
            f.Project.References.Items.Add(new FakeReference { GUID = "{CCCCCCCC-CCCC-CCCC-CCCC-CCCCCCCCCCCC}", Major = 3, Minor = 0, ThrowMetadata = true });
            dynamic inaccessible = f.Session.Execute(new Request { Command = "list_references", Project = f.Project.Name }).Data;
            var inaccessibleEntries = ((IEnumerable)inaccessible.References).Cast<object>().ToArray();
            Assert.IsNull(inaccessibleEntries[2].GetType().GetProperty("Name").GetValue(inaccessibleEntries[2]));
            Assert.IsNull(inaccessibleEntries[2].GetType().GetProperty("FullPath").GetValue(inaccessibleEntries[2]));
        }

        [TestMethod]
        public void AddReferenceGuidRequiresFreshVersionAndRejectsDuplicateIdentity()
        {
            var f = Create();
            const string guid = "{A04D9E3D-48C7-4B46-A582-F00DD08E89CA}";
            var request = new Request
            {
                Command = "add_reference_guid",
                Project = f.Project.Name,
                Guid = guid,
                Major = 1,
                Minor = 0
            };
            Assert.ThrowsException<ArgumentException>(() => f.Session.Execute(request));
            request.ExpectedReferencesVersion = (string)((dynamic)f.Session.Execute(new Request { Command = "list_references", Project = f.Project.Name }).Data).Version;
            request.Guid = "invalid";
            Assert.ThrowsException<ArgumentException>(() => f.Session.Execute(request));
            request.Guid = guid;
            request.ExpectedReferencesVersion = Sha("stale");
            Assert.ThrowsException<InvalidOperationException>(() => f.Session.Execute(request));
            request.ExpectedReferencesVersion = (string)((dynamic)f.Session.Execute(new Request { Command = "list_references", Project = f.Project.Name }).Data).Version;
            f.Project.Mode = 1;
            Assert.ThrowsException<InvalidOperationException>(() => f.Session.Execute(request));
            f.Project.Mode = 2;
            var added = f.Session.Execute(request);
            Assert.IsTrue(added.Ok);
            Assert.AreEqual(1, f.Project.References.Items.Count);
            request.ExpectedReferencesVersion = (string)((dynamic)added.Data).Version;
            Assert.ThrowsException<InvalidOperationException>(() => f.Session.Execute(request));
            Assert.AreEqual(1, f.Project.References.Items.Count);
        }

        [TestMethod]
        public void RemoveReferenceRequiresExactVersionAndRefusesBuiltInReference()
        {
            var f = Create();
            const string guid = "{D9C22780-B36A-4F27-8B07-29FC38744610}";
            var builtIn = new FakeReference
            {
                GUID = guid,
                Major = 1,
                Minor = 2,
                Name = "BuiltIn",
                BuiltIn = true
            };
            f.Project.References.Items.Add(builtIn);
            var request = new Request
            {
                Command = "remove_reference",
                Project = f.Project.Name,
                Guid = guid,
                Major = 1,
                Minor = 2,
                ExpectedReferencesVersion = (string)((dynamic)f.Session.Execute(new Request { Command = "list_references", Project = f.Project.Name }).Data).Version
            };
            Assert.ThrowsException<InvalidOperationException>(() => f.Session.Execute(request));
            request.Minor = 3;
            Assert.ThrowsException<InvalidOperationException>(() => f.Session.Execute(request));
            request.Minor = 2;
            builtIn.BuiltIn = false;
            Assert.ThrowsException<InvalidOperationException>(() => f.Session.Execute(request));
            request.ExpectedReferencesVersion = (string)((dynamic)f.Session.Execute(new Request { Command = "list_references", Project = f.Project.Name }).Data).Version;
            Assert.IsTrue(f.Session.Execute(request).Ok);
            Assert.AreEqual(0, f.Project.References.Items.Count);
        }

        [TestMethod]
        public void AddReferenceFileRequiresExistingAbsolutePathAndFreshVersion()
        {
            var f = Create();
            var request = new Request
            {
                Command = "add_reference_file",
                Project = f.Project.Name,
                Path = "relative.dll"
            };
            Assert.ThrowsException<ArgumentException>(() => f.Session.Execute(request));
            request.Path = @"C:\DefinitelyMissing\CodexVBE-test.tlb";
            Assert.ThrowsException<FileNotFoundException>(() => f.Session.Execute(request));
            string path = Path.GetTempFileName();
            try
            {
                request.Path = path;
                request.ExpectedReferencesVersion = (string)((dynamic)f.Session.Execute(new Request { Command = "list_references", Project = f.Project.Name }).Data).Version;
                var result = f.Session.Execute(request);
                Assert.IsTrue(result.Ok);
                Assert.AreEqual(Path.GetFullPath(path), f.Project.References.Items[0].FullPath);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [TestMethod]
        public void GitScopeRequiresSavedAbsoluteHostPath()
        {
            var f = Create();
            Assert.AreEqual(Path.GetFullPath(f.Project.FileName), f.Session.GitScope(f.Project.Name));
            f.Project.FileName = string.Empty;
            Assert.ThrowsException<InvalidOperationException>(() => f.Session.GitScope(f.Project.Name));
        }

        [TestMethod]
        public void SigningRejectsMissingTokenStaleProjectRunningModeUnsavedAndBadThumbprint()
        {
            var f = Create();
            var request = new Request
            {
                Command = "sign_project",
                Project = f.Project.Name,
                ExpectedMode = 2,
                CertificateThumbprint = "invalid"
            };
            Assert.ThrowsException<ArgumentException>(() => f.Session.Execute(request));
            request.ExpectedProjectVersion = "stale";
            f.Project.Mode = 1;
            Assert.ThrowsException<InvalidOperationException>(() => f.Session.Execute(request));
            f.Project.Mode = 2;
            Assert.ThrowsException<InvalidOperationException>(() => f.Session.Execute(request));
            dynamic state = f.Session.Execute(new Request { Command = "project_properties", Project = f.Project.Name }).Data;
            request.ExpectedProjectVersion = state.Version;
            f.Project.Saved = false;
            Assert.ThrowsException<InvalidOperationException>(() => f.Session.Execute(request));
            f.Project.Saved = true;
            Assert.ThrowsException<ArgumentException>(() => f.Session.Execute(request));
        }

        [TestMethod]
        public void ProjectDispatchReturnsVersionPersistenceAndSignatureHostLimits()
        {
            var f = Create();
            dynamic properties = f.Session.Execute(new Request { Command = "project_properties", Project = f.Project.Name }).Data;
            Assert.AreEqual(f.Project.Name, (string)properties.Project);
            Assert.AreEqual(2, (int)properties.Mode);
            Assert.IsFalse(string.IsNullOrWhiteSpace((string)properties.Version));
            Assert.AreEqual(1, (int)properties.Components.Count);
            f.Project.Mode = 1;
            dynamic changed = f.Session.Execute(new Request { Command = "project_properties", Project = f.Project.Name }).Data;
            Assert.AreNotEqual((string)properties.Version, (string)changed.Version);
            f.Project.Saved = false;
            dynamic persistence = f.Session.Execute(new Request { Command = "project_persistence_status", Project = f.Project.Name }).Data;
            Assert.IsFalse((bool)persistence.ProjectSaved);
            Assert.IsFalse((bool)persistence.HostAvailable);
            Assert.IsNull((object)persistence.HostPath);
            dynamic signature = f.Session.Execute(new Request { Command = "project_signature_status", Project = f.Project.Name }).Data;
            Assert.IsFalse((bool)signature.Available);
            Assert.IsNull((object)signature.Signed);
            Assert.AreEqual("Host", (string)signature.Source);
            dynamic persistSignature = f.Session.PersistProjectSignature(f.Project.Name);
            Assert.IsFalse((bool)persistSignature.Available);
            Assert.IsFalse((bool)persistSignature.Saved);
        }

        [TestMethod]
        public void ProjectMutationDispatchRejectsUnsupportedHostAndDisabledRename()
        {
            var f = Create();
            var rename = f.Session.Execute(new Request { Command = "rename_project", Project = f.Project.Name, NewName = "Other" });
            Assert.IsFalse(rename.Ok);
            StringAssert.Contains(rename.Error, "disabled");
            Assert.AreEqual("VBAProject", f.Project.Name);
            Assert.ThrowsException<ArgumentException>(() => f.Session.Execute(new Request { Command = "save_host_document", Project = f.Project.Name }));
            Assert.ThrowsException<InvalidOperationException>(() => f.Session.Execute(new Request { Command = "save_host_document", Project = f.Project.Name, ExpectedHostPath = @"C:\Temp\Host.xlsm" }));
            Assert.ThrowsException<ArgumentException>(() => f.Session.Execute(new Request { Command = "save_host_document_as", Project = f.Project.Name }));
            Assert.ThrowsException<InvalidOperationException>(() => f.Session.Execute(new Request { Command = "save_host_document_as", Project = f.Project.Name, Path = @"C:\Temp\Other.xlsm", ExpectedProjectVersion = "stale" }));
        }

        [TestMethod]
        public void SignatureDialogDispatchRejectsMissingModeInactiveAndRunningProject()
        {
            var f = Create();
            Assert.ThrowsException<ArgumentException>(() => f.Session.Execute(new Request { Command = "read_project_signature_dialog", Project = f.Project.Name }));
            f.Project.Mode = 1;
            Assert.ThrowsException<InvalidOperationException>(() => f.Session.Execute(new Request { Command = "read_project_signature_dialog", Project = f.Project.Name, ExpectedMode = 2 }));
            f.Project.Mode = 2;
            Assert.ThrowsException<InvalidOperationException>(() => f.Session.Execute(new Request { Command = "read_project_signature_dialog", Project = f.Project.Name, ExpectedMode = 2 }));
        }

        [TestMethod]
        public void ReferenceMutationRejectsNegativeVersionsMissingIdentityAndChangedInventory()
        {
            var f = Create();
            const string guid = "{A04D9E3D-48C7-4B46-A582-F00DD08E89CA}";
            string version = (string)((dynamic)f.Session.Execute(new Request { Command = "list_references", Project = f.Project.Name }).Data).Version;
            foreach (string command in new[]
            {
                "add_reference_guid",
                "remove_reference"
            }

            )
            {
                var invalid = new Request
                {
                    Command = command,
                    Project = f.Project.Name,
                    Guid = guid,
                    Major = -1,
                    Minor = 0,
                    ExpectedReferencesVersion = version
                };
                Assert.ThrowsException<ArgumentException>(() => f.Session.Execute(invalid));
                invalid.Major = 0;
                invalid.Minor = -1;
                Assert.ThrowsException<ArgumentException>(() => f.Session.Execute(invalid));
                invalid.Minor = 0;
                if (command == "add_reference_guid")
                    invalid.ExpectedReferencesVersion = "stale";
                Assert.ThrowsException<InvalidOperationException>(() => f.Session.Execute(invalid));
            }

            var add = new Request
            {
                Command = "add_reference_guid",
                Project = f.Project.Name,
                Guid = guid.ToLowerInvariant(),
                Major = 1,
                Minor = 0,
                ExpectedReferencesVersion = version.ToUpperInvariant()
            };
            Assert.IsTrue(f.Session.Execute(add).Ok);
            Assert.AreEqual(1, f.Project.References.Items.Count);
            var remove = new Request
            {
                Command = "remove_reference",
                Project = f.Project.Name,
                Guid = guid,
                Major = 1,
                Minor = 0,
                ExpectedReferencesVersion = version
            };
            Assert.ThrowsException<InvalidOperationException>(() => f.Session.Execute(remove));
            remove.ExpectedReferencesVersion = (string)((dynamic)f.Session.Execute(new Request { Command = "list_references", Project = f.Project.Name }).Data).Version;
            Assert.IsTrue(f.Session.Execute(remove).Ok);
            Assert.AreEqual(0, f.Project.References.Items.Count);
        }
    }
}
