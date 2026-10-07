using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace VBAi.Tests.Unit
{
    public sealed partial class ToolbarProfilesTests
    {
        [TestMethod]
        public void SessionRestoresOnlyKnownHostProfilesFromTheirOwnTemporaryDatabase()
        {
            using (var scope = new ProfileScope())
            {
                foreach (string hostName in new[] { "excel", "sLdWoRkS" })
                {
                    string path = System.IO.Path.Combine(scope.Root, "VBAi", "VbeToolbars", hostName.ToUpperInvariant() + ".sqlite");
                    var profile = ProfileBar(); var profiles = new VBAi.VbeToolbarProfiles(path); profiles.Update(profile.Name, profile);
                    var host = NativeHost(); int paths = 0;
                    var session = new VBAi.VbeSession(host, null, null, () => hostName, () => { paths++; return scope.Root; });
                    Assert.AreEqual(1, paths); Assert.AreEqual(2, host.CommandBars.Count); Assert.AreEqual(profile.Name, host.CommandBars[1].Name);
                    Assert.AreEqual(profile.Commands[0].Tag, host.CommandBars[1].Controls[1].Tag);
                    var response = session.Execute(new VBAi.Request { Command = "list_toolbars" }); Assert.IsTrue(response.Ok, response.Error);
                    Assert.AreEqual(0, ((System.Collections.IList)Data(response.Data)["ProfileErrors"]).Count);
                }
                var unknownHost = NativeHost();
                var unknown = new VBAi.VbeSession(unknownHost, null, null, () => "OTHER", () => { Assert.Fail("Other processes must not read profile storage."); return scope.Root; });
                Assert.AreEqual(1, unknownHost.CommandBars.Count); Assert.IsTrue(unknown.Execute(new VBAi.Request { Command = "list_toolbars" }).Ok);
                Assert.AreEqual(2, System.IO.Directory.GetFiles(System.IO.Path.Combine(scope.Root, "VBAi", "VbeToolbars"), "*.sqlite").Length);
            }
        }

        [TestMethod]
        public void SessionStartsWithEmptyProfilesAndReportsCorruptionWithoutDroppingServices()
        {
            using (var scope = new ProfileScope())
            {
                var host = NativeHost(); var empty = new VBAi.VbeSession(host, null, null, () => "EXCEL", () => scope.Root);
                Assert.AreEqual(1, host.CommandBars.Count); Assert.IsTrue(empty.Execute(new VBAi.Request { Command = "list_toolbars" }).Ok);
                string path = System.IO.Path.Combine(scope.Root, "VBAi", "VbeToolbars", "EXCEL.sqlite");
                using (var store = new VBAi.ChatSessionStore(path)) store.UpdateToolbarProfile("Invalid", new VBAi.VbeToolbarProfiles.Bar { Name = "Invalid", Commands = new VBAi.VbeToolbarProfiles.Command[0] }, bars => { });
                var corrupt = new VBAi.VbeSession(host, null, null, () => "EXCEL", () => scope.Root);
                var response = corrupt.Execute(new VBAi.Request { Command = "list_toolbars" }); Assert.IsTrue(response.Ok, response.Error);
                var errors = (System.Collections.IList)Data(response.Data)["ProfileErrors"]; Assert.AreEqual(1, errors.Count);
                StringAssert.Contains((string)errors[0], "Invalid toolbar profile contents."); Assert.AreEqual(1, host.CommandBars.Count);
            }
        }
    }
}
namespace VBAi.Tests.Unit
{
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using VBAi;

    [TestClass]
    [TestCategory("Unit")]
    public sealed partial class VbeSessionContractTests
    {
        [TestMethod]
        public void GeneralQuarantineRefusesInventoryAndAllNativeAsyncRoutes()
        {
            var session = new VbeSession(new FakeVbe());
            typeof(VbeSession).GetField("generalQuarantined", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic).SetValue(session, true);
            Assert.IsTrue(session.Execute(new Request { Command = "status" }).Ok);
            foreach (string command in new[] { "list_projects", "project_properties", "set_project_property", "open_native_ide_dialog" })
                Assert.IsFalse(session.Execute(new Request { Command = command }).Ok, command);
            Assert.ThrowsException<InvalidOperationException>(() => session.SaveHostDocumentAsync(new Request()));
            Assert.ThrowsException<InvalidOperationException>(() => session.ReadImmediateAsync(new Request()));
            Assert.ThrowsException<InvalidOperationException>(() => session.InspectLocalScalarsAsync(new Request()));
        }

        [TestMethod]
        public void GeneralPendingAllowsOnlyLiveAuthorizationInventoryAndRejectsSyncGeneral()
        {
            var session = new VbeSession(new FakeVbe());
            var flags = System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic;
            Assert.IsFalse(session.Execute(new Request { Command = "read_project_general" }).Ok);
            Assert.IsFalse(session.Execute(new Request { Command = "set_project_general" }).Ok);
            typeof(VbeSession).GetField("generalInFlight", flags).SetValue(session, true);
            Assert.IsFalse(session.Execute(new Request { Command = "list_projects" }).Ok);
            typeof(VbeSession).GetField("generalAuthorizationDepth", flags).SetValue(session, 1);
            Assert.IsTrue(session.Execute(new Request { Command = "list_projects" }).Ok);
            Assert.IsFalse(session.Execute(new Request { Command = "set_project_property" }).Ok);
            typeof(VbeSession).GetField("generalQuarantined", flags).SetValue(session, true);
            Assert.IsFalse(session.Execute(new Request { Command = "list_projects" }).Ok);
        }

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
        public void EditorModuleResolutionUsesExactProjectAndCaseInsensitiveComponentIdentity()
        {
            var f = new VBAi.Tests.Infrastructure.EditorVbeContract(); var session = new VbeSession(f.Vbe);
            Assert.AreSame(f.Original, ((EditorVbeModule)session.ResolveEditorModule("project1", "module1")).Component);
            Assert.ThrowsException<ArgumentException>(() => session.ResolveEditorModule("Project1", null));
            Assert.ThrowsException<ArgumentException>(() => session.ResolveEditorModule("Project1", " "));
            Assert.ThrowsException<InvalidOperationException>(() => session.ResolveEditorModule("Project1", "Missing"));
            Assert.IsNull(session.ReferenceEventSource(null)); Assert.IsNull(session.ReferenceEventSource(""));
        }

        [TestMethod]
        public void RecentIdeRoutesReachTheirServiceGuardsAndAsyncOnlyResponses()
        {
            var f = new VBAi.Tests.Infrastructure.EditorVbeContract(); var session = new VbeSession(f.Vbe);
            foreach (string command in new[] { "preview_procedure_rename", "apply_procedure_rename", "preview_class_member_rename", "apply_class_member_rename", "open_native_ide_dialog", "read_project_protection", "set_project_protection", "project_collection_state", "create_standalone_project", "open_standalone_project", "close_standalone_project", "open_project_help", "list_macros", "read_navigation_surface", "change_navigation_surface", "run_procedure_values", "procedure_values_status", "preview_fit_form_content", "apply_fit_form_content" })
            {
                try
                {
                    var response = session.Execute(new Request { Command = command });
                    Assert.IsNotNull(response, command);
                    Assert.IsFalse((response.Error ?? "").StartsWith("Unknown command:"), command);
                    if (command == "read_project_protection" || command == "set_project_protection") Assert.AreEqual("ExpectedMode=2 is required for project properties.", response.Error);
                    if (command == "read_navigation_surface" || command == "change_navigation_surface") StringAssert.Contains(response.Error, "requires InvokeAsync");
                }
                catch (ArgumentException) { }
                catch (InvalidOperationException) { }
            }
            foreach (string command in new[] { "read_project_protection", "set_project_protection" })
            {
                Assert.ThrowsException<ArgumentException>(() => session.Execute(new Request { Command = command, ExpectedMode = 2 }));
            }
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

namespace VBAi.Tests.Unit
{
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using System;
    using System.Collections;
    using System.Linq;
    using System.Reflection;
    using System.Security.Cryptography.X509Certificates;
    using VBAi;

    [TestClass]
    [TestCategory("Unit")]
    public sealed partial class VbeSessionCoverageTests
    {
        [TestMethod]
        public void SigningRequiresAllIdentityFieldsBeforeReadingAnyCertificateStore()
        {
            using (var fixture = new SigningFixture())
            {
                var method = typeof(VbeSession).GetMethod("BeginSignProject", BindingFlags.Instance | BindingFlags.NonPublic);
                var nullRequest = Assert.ThrowsException<TargetInvocationException>(() => method.Invoke(fixture.Session, new object[] { null }));
                Assert.IsInstanceOfType(nullRequest.InnerException, typeof(ArgumentException));
                foreach (var field in new[] { "Project", "ExpectedProjectVersion", "CertificateThumbprint" })
                    foreach (var value in new[] { null, "", " " })
                    {
                        var request = fixture.Request(new string('A', 40));
                        typeof(Request).GetProperty(field).SetValue(request, value);
                        var error = Assert.ThrowsException<ArgumentException>(() => fixture.Session.Execute(request));
                        StringAssert.Contains(error.Message, "are required");
                    }
                Assert.AreEqual(0, fixture.Opened.Count);
                Assert.AreEqual(0, fixture.Scheduled);
            }
        }

        [TestMethod]
        public void SigningRejectsChangedModeRevisionUnsavedProjectAndInvalidThumbprint()
        {
            using (var fixture = new SigningFixture())
            {
                var request = fixture.Request(new string('A', 40));
                request.ExpectedMode = 1;
                StringAssert.Contains(Assert.ThrowsException<InvalidOperationException>(() => fixture.Session.Execute(request)).Message, "design mode");
                fixture.Project.Mode = 1;
                request = fixture.Request(new string('A', 40));
                StringAssert.Contains(Assert.ThrowsException<InvalidOperationException>(() => fixture.Session.Execute(request)).Message, "design mode");
                fixture.Project.Mode = 2;
                request = fixture.Request(new string('A', 40));
                request.ExpectedProjectVersion = "stale";
                StringAssert.Contains(Assert.ThrowsException<InvalidOperationException>(() => fixture.Session.Execute(request)).Message, "changed");
                fixture.Project.Saved = false;
                request = fixture.Request(new string('A', 40));
                StringAssert.Contains(Assert.ThrowsException<InvalidOperationException>(() => fixture.Session.Execute(request)).Message, "Save the VBA project");
                fixture.Project.Saved = true;
                foreach (var thumbprint in new[] { "invalid", new string('A', 39), new string('G', 40) })
                    Assert.ThrowsException<ArgumentException>(() => fixture.Session.Execute(fixture.Request(thumbprint)));
                Assert.AreEqual(0, fixture.Opened.Count);
                Assert.AreEqual(0, fixture.Scheduled);
            }
        }

        [TestMethod]
        public void SigningRejectsMissingDuplicateKeylessExpiredFutureAndNonSigningCertificates()
        {
            using (var fixture = new SigningFixture())
            using (var valid = new CertificateFixture())
            using (var publicOnly = new X509Certificate2(valid.Certificate.Export(X509ContentType.Cert)))
            using (var expired = new CertificateFixture("Expired", true, after: new DateTimeOffset(2029, 2, 1, 0, 0, 0, TimeSpan.Zero)))
            using (var future = new CertificateFixture("Future", true, before: new DateTimeOffset(2030, 2, 1, 0, 0, 0, TimeSpan.Zero)))
            using (var wrongUsage = new CertificateFixture("No code signing", false))
            {
                var missing = Assert.ThrowsException<InvalidOperationException>(() => fixture.Session.Execute(fixture.Request(valid.Certificate.Thumbprint)));
                StringAssert.Contains(missing.Message, "absent or ambiguous");
                fixture.Personal.Certificates.Add(valid.Certificate);
                fixture.Personal.Certificates.Add(valid.Certificate);
                StringAssert.Contains(Assert.ThrowsException<InvalidOperationException>(() => fixture.Session.Execute(fixture.Request(valid.Certificate.Thumbprint))).Message, "absent or ambiguous");
                foreach (var certificate in new[] { publicOnly, expired.Certificate, future.Certificate, wrongUsage.Certificate })
                {
                    fixture.Personal.Certificates.Clear();
                    fixture.Personal.Certificates.Add(certificate);
                    var error = Assert.ThrowsException<InvalidOperationException>(() => fixture.Session.Execute(fixture.Request(certificate.Thumbprint)));
                    StringAssert.Contains(error.Message, certificate == wrongUsage.Certificate ? "code signing" : "private key and current validity");
                }
                Assert.AreEqual(0, fixture.Scheduled);
                Assert.IsTrue(fixture.Personal.Disposed);
                Assert.IsNull(fixture.Machine.Flags);
            }
        }

        [TestMethod]
        public void SigningRejectsEmptyAndAmbiguousCertificateDisplayNamesAcrossBothStores()
        {
            using (var fixture = new SigningFixture())
            using (var selected = new CertificateFixture("Same name"))
            using (var collision = new CertificateFixture("Same name"))
            using (var different = new CertificateFixture("Different name"))
            using (var nameless = new CertificateFixture(""))
            {
                fixture.Personal.Certificates.Add(nameless.Certificate);
                StringAssert.Contains(Assert.ThrowsException<InvalidOperationException>(() => fixture.Session.Execute(fixture.Request(nameless.Certificate.Thumbprint))).Message, "display name is empty");
                fixture.Personal.Certificates.Clear();
                fixture.Personal.Certificates.Add(selected.Certificate);
                fixture.Personal.Certificates.Add(different.Certificate);
                fixture.Machine.Certificates.Add(collision.Certificate);
                StringAssert.Contains(Assert.ThrowsException<InvalidOperationException>(() => fixture.Session.Execute(fixture.Request(selected.Certificate.Thumbprint))).Message, "ambiguous across");
                Assert.AreEqual(0, fixture.Scheduled);
                Assert.IsTrue(fixture.Personal.Disposed);
                Assert.IsTrue(fixture.Machine.Disposed);
                fixture.Machine.Certificates.Clear();
                fixture.Personal.Certificates.Add(collision.Certificate);
                StringAssert.Contains(Assert.ThrowsException<InvalidOperationException>(() => fixture.Session.Execute(fixture.Request(selected.Certificate.Thumbprint))).Message, "ambiguous across");
            }
        }

        [TestMethod]
        public void SigningSchedulesTheExactCertificateOnlyAfterValidationAndDisposesBothStores()
        {
            using (var fixture = new SigningFixture())
            using (var selected = new CertificateFixture())
            using (var other = new CertificateFixture("Other certificate"))
            {
                fixture.Personal.Certificates.Add(selected.Certificate);
                fixture.Personal.Certificates.Add(other.Certificate);
                fixture.Machine.Certificates.Add(other.Certificate);
                fixture.Machine.Certificates.Add(selected.Certificate);
                for (int attempt = 0; attempt < 2; attempt++)
                {
                    var lower = selected.Certificate.Thumbprint.ToLowerInvariant();
                    var spaced = string.Join(" ", Enumerable.Range(0, lower.Length / 2).Select(i => lower.Substring(i * 2, 2)));
                    var request = fixture.Request(spaced);
                    request.ExpectedProjectVersion = request.ExpectedProjectVersion.ToUpperInvariant();
                    dynamic result = fixture.Session.Execute(request).Data;
                    Assert.IsTrue((bool)result.Scheduled);
                    Assert.IsFalse((bool)result.UnsignedVerified);
                    Assert.AreEqual(fixture.Project.Name, (string)result.Project);
                    Assert.AreEqual(selected.Certificate.Thumbprint, (string)result.CertificateThumbprint);
                    Assert.AreEqual("VBAi Unit Signing", (string)result.CertificateName);
                    Assert.AreSame(fixture.NativeResult, (object)result.NativeCommand);
                    Assert.AreSame(request, fixture.ScheduledRequest);
                }
                Assert.AreEqual(2, fixture.Scheduled);
                Assert.AreEqual(OpenFlags.ReadOnly, fixture.Personal.Flags);
                Assert.AreEqual(OpenFlags.ReadOnly, fixture.Machine.Flags);
                Assert.IsTrue(fixture.Personal.Disposed);
                Assert.IsTrue(fixture.Machine.Disposed);
                Assert.IsTrue(fixture.Project.Saved);
                Assert.IsFalse(fixture.Workbook.VBASigned);
            }
        }

        [TestMethod]
        public void SigningPropagatesStoreAndSchedulerFailuresWithoutLeakingOpenedStores()
        {
            using (var fixture = new SigningFixture())
            using (var selected = new CertificateFixture())
            {
                fixture.Personal.Certificates.Add(selected.Certificate);
                var personalFailure = new InvalidOperationException("personal store denied");
                fixture.Personal.OpenError = personalFailure;
                Assert.AreSame(personalFailure, Assert.ThrowsException<InvalidOperationException>(() => fixture.Session.Execute(fixture.Request(selected.Certificate.Thumbprint))));
                Assert.IsTrue(fixture.Personal.Disposed);
                fixture.Personal.OpenError = null;
                var machineFailure = new InvalidOperationException("machine store denied");
                fixture.Machine.OpenError = machineFailure;
                Assert.AreSame(machineFailure, Assert.ThrowsException<InvalidOperationException>(() => fixture.Session.Execute(fixture.Request(selected.Certificate.Thumbprint))));
                Assert.IsTrue(fixture.Machine.Disposed);
                fixture.Machine.OpenError = null;
                var schedulerFailure = new InvalidOperationException("native queue denied");
                fixture.Session.SignatureScheduler = request => { throw schedulerFailure; };
                Assert.AreSame(schedulerFailure, Assert.ThrowsException<InvalidOperationException>(() => fixture.Session.Execute(fixture.Request(selected.Certificate.Thumbprint))));
                Assert.IsTrue(fixture.Personal.Disposed);
                Assert.IsTrue(fixture.Machine.Disposed);
                Assert.AreEqual(0, fixture.Scheduled);
            }
        }

        [TestMethod]
        public void ExcelSigningRequiresUnsignedSavedMacroFileAndRealVbaContent()
        {
            using (var fixture = new SigningFixture(true))
            using (var selected = new CertificateFixture())
            {
                fixture.Personal.Certificates.Add(selected.Certificate);
                fixture.Workbook.VBASigned = true;
                StringAssert.Contains(Assert.ThrowsException<InvalidOperationException>(() => fixture.Session.Execute(fixture.Request(selected.Certificate.Thumbprint))).Message, "existing VBA signature");
                fixture.Workbook.VBASigned = false;
                foreach (var path in new[] { null, "", "relative.xlsm", fixture.FilePath + ".missing" })
                {
                    fixture.SetPath(path);
                    StringAssert.Contains(Assert.ThrowsException<InvalidOperationException>(() => fixture.Session.Execute(fixture.Request(selected.Certificate.Thumbprint))).Message, "Save the macro-enabled");
                }
                fixture.SetPath(fixture.FilePath);
                fixture.Project.ThrowFileName = true;
                StringAssert.Contains(Assert.ThrowsException<InvalidOperationException>(() => fixture.Session.Execute(fixture.Request(selected.Certificate.Thumbprint))).Message, "Save the macro-enabled");
                fixture.Project.ThrowFileName = false;
                fixture.Host.IsExcel = false;
                StringAssert.Contains(Assert.ThrowsException<InvalidOperationException>(() => fixture.Session.Execute(fixture.Request(selected.Certificate.Thumbprint))).Message, "Save the macro-enabled");
                fixture.Host.IsExcel = true;
                StringAssert.Contains(Assert.ThrowsException<InvalidOperationException>(() => fixture.Session.Execute(fixture.Request(selected.Certificate.Thumbprint))).Message, "no VBA content");
                var document = new VbeSessionTests.FakeComponent { Name = "ThisWorkbook", Type = 100, CodeModule = new VbeSessionTests.FakeModule("") };
                fixture.Project.VBComponents.Items.Add(document);
                StringAssert.Contains(Assert.ThrowsException<InvalidOperationException>(() => fixture.Session.Execute(fixture.Request(selected.Certificate.Thumbprint))).Message, "no VBA content");
                document.CodeModule = new VbeSessionTests.FakeModule("Option Explicit");
                dynamic documentResult = fixture.Session.Execute(fixture.Request(selected.Certificate.Thumbprint)).Data;
                Assert.IsTrue((bool)documentResult.UnsignedVerified);
                document.CodeModule = new VbeSessionTests.FakeModule("");
                fixture.Project.VBComponents.Items.Add(new VbeSessionTests.FakeComponent { Name = "Module1", Type = 1, CodeModule = new VbeSessionTests.FakeModule("") });
                Assert.IsTrue(fixture.Session.Execute(fixture.Request(selected.Certificate.Thumbprint)).Ok);
                Assert.AreEqual(2, fixture.Scheduled);
                Assert.IsFalse(fixture.Workbook.VBASigned);
            }
        }

        [TestMethod]
        public void ExcelSigningRejectsAnExistingNonMacroFormatBeforeReadingCertificates()
        {
            using (var fixture = new SigningFixture(true))
            {
                string path = fixture.FilePath + ".txt";
                System.IO.File.WriteAllText(path, "format fixture");
                try
                {
                    fixture.SetPath(path);
                    StringAssert.Contains(Assert.ThrowsException<InvalidOperationException>(() => fixture.Session.Execute(fixture.Request(new string('A', 40)))).Message, "format does not support");
                    Assert.AreEqual(0, fixture.Opened.Count);
                }
                finally { System.IO.File.Delete(path); }
            }
        }

        [TestMethod]
        public void CertificateCatalogueFiltersUsageAndPrivateKeyAndReportsValidityWithoutSigning()
        {
            using (var fixture = new SigningFixture())
            using (var valid = new CertificateFixture())
            using (var noUsage = new CertificateFixture("No usage", false))
            using (var expired = new CertificateFixture("Expired", true, after: new DateTimeOffset(2029, 2, 1, 0, 0, 0, TimeSpan.Zero)))
            using (var future = new CertificateFixture("Future", true, before: new DateTimeOffset(2030, 2, 1, 0, 0, 0, TimeSpan.Zero)))
            using (var publicOnly = new X509Certificate2(valid.Certificate.Export(X509ContentType.Cert)))
            {
                fixture.Personal.Certificates.AddRange(new[] { publicOnly, noUsage.Certificate, valid.Certificate, expired.Certificate, future.Certificate });
                var rows = ((IEnumerable)fixture.Session.Execute(new Request { Command = "list_signing_certificates" }).Data).Cast<dynamic>().ToArray();
                Assert.AreEqual(3, rows.Length);
                Assert.AreEqual(valid.Certificate.Thumbprint, (string)rows[0].Thumbprint);
                Assert.IsTrue((bool)rows[0].EligibleNow);
                Assert.IsFalse((bool)rows[1].EligibleNow);
                Assert.IsFalse((bool)rows[2].EligibleNow);
                Assert.AreEqual(valid.Certificate.Subject, (string)rows[0].Subject);
                Assert.AreEqual(valid.Certificate.Issuer, (string)rows[0].Issuer);
                Assert.AreEqual(valid.Certificate.NotBefore.ToString("o"), (string)rows[0].NotBefore);
                Assert.AreEqual(valid.Certificate.NotAfter.ToString("o"), (string)rows[0].NotAfter);
                Assert.IsTrue(fixture.Personal.Disposed);
                Assert.AreEqual(0, fixture.Scheduled);
            }
        }

        [TestMethod]
        public void DefaultSigningServicesReadOnlyStoresAndUseTheActualProcessAndClock()
        {
            var session = new VbeSession(new SessionHost());
            Assert.AreEqual(System.Diagnostics.Process.GetCurrentProcess().ProcessName, session.SigningProcessName());
            Assert.IsTrue(Math.Abs((DateTime.Now - session.SigningClock()).TotalSeconds) < 5);
            foreach (var location in new[] { StoreLocation.CurrentUser, StoreLocation.LocalMachine })
                using (var store = session.SigningStore(location))
                {
                    store.Open(OpenFlags.ReadOnly);
                    Assert.IsNotNull(store.Certificates);
                }
        }

        [TestMethod]
        public void DefaultSchedulerRetainsTheNativeExactProjectCheck()
        {
            using (var fixture = new SigningFixture())
            using (var selected = new CertificateFixture())
            {
                var session = new VbeSession(fixture.Vbe, fixture.Host);
                session.SigningClock = fixture.Session.SigningClock;
                session.SigningStore = fixture.Session.SigningStore;
                session.SigningProcessName = fixture.Session.SigningProcessName;
                fixture.Personal.Certificates.Add(selected.Certificate);
                fixture.Vbe.ActiveVBProject = null;
                StringAssert.Contains(Assert.ThrowsException<InvalidOperationException>(() => session.Execute(fixture.Request(selected.Certificate.Thumbprint))).Message, "Select the exact project");
                Assert.IsTrue(fixture.Personal.Disposed);
                Assert.IsTrue(fixture.Machine.Disposed);
            }
        }

        [TestMethod]
        public void GitProjectResolvesTheSelectedLiveProjectAndRefusesChangedDocumentPath()
        {
            using (var fixture = new SigningFixture())
            {
                var linked = fixture.Session.GitProject(fixture.Project.Name, fixture.FilePath + ".other");
                var error = Assert.ThrowsException<InvalidOperationException>(() => linked.Capture());
                StringAssert.Contains(error.Message, UiText.Get("The linked document changed. Reopen GitHub integration."));
                Assert.IsTrue(fixture.Project.Saved);
            }
        }

        [TestMethod]
        public void ComponentCreationRefusesWrongNativeNameOrTypeAndRetainsOriginalRenameFailure()
        {
            foreach (var wrongName in new[] { true, false })
            {
                var project = new VbeSessionTests.FakeProject { Name = "Components", Mode = 2 };
                var host = new SessionHost();
                host.VBProjects.Add(project);
                if (wrongName) project.VBComponents.ForcedReadbackName = "Unexpected";
                else project.VBComponents.ForcedAddedType = 3;
                var session = new VbeSession(host);
                StringAssert.Contains(Assert.ThrowsException<InvalidOperationException>(() => session.Execute(new Request
                {
                    Command = "create_module",
                    Project = project.Name,
                    Module = "Created",
                    ExpectedMode = 2
                })).Message,
                    "did not create the requested component identity");
                Assert.AreEqual(1, project.VBComponents.Items.Count);
            }
            var rejected = new VbeSessionTests.FakeProject { Name = "Rejected", Mode = 2 };
            rejected.VBComponents.RejectedName = "RejectedModule";
            rejected.VBComponents.FailRemoval = true;
            var rejectingHost = new SessionHost();
            rejectingHost.VBProjects.Add(rejected);
            var rejectingSession = new VbeSession(rejectingHost);
            var failure = Assert.ThrowsException<InvalidOperationException>(() => rejectingSession.Execute(new Request
            {
                Command = "create_class",
                Project = rejected.Name,
                Module = "RejectedModule",
                ExpectedMode = 2
            }));
            Assert.AreEqual("Rejected by VBE", failure.Message);
            Assert.AreEqual(1, rejected.VBComponents.RemoveCount);
            Assert.AreEqual(1, rejected.VBComponents.Items.Count);
        }

        [TestMethod]
        public void ComponentCreationRejectsAProjectThatHasLeftDesignModeBeforeAddingAnything()
        {
            foreach (var command in new[] { "create_module", "create_class" })
            {
                var project = new VbeSessionTests.FakeProject { Name = "Executing", Mode = 1 };
                var host = new SessionHost();
                host.VBProjects.Add(project);
                var session = new VbeSession(host);
                var failure = Assert.ThrowsException<InvalidOperationException>(() => session.Execute(new Request
                {
                    Command = command,
                    Project = project.Name,
                    Module = "NewComponent",
                    ExpectedMode = 2
                }));
                Assert.AreEqual("The project is no longer in design mode.", failure.Message);
                Assert.AreEqual(0, project.VBComponents.Items.Count);
            }
        }

        [TestMethod]
        public void ReferenceFileAndRemovalRejectMissingPathsAndEveryIdentityMismatchBeforeMutation()
        {
            var project = new VbeSessionTests.FakeProject { Name = "References", Mode = 2 };
            var host = new SessionHost();
            host.VBProjects.Add(project);
            var session = new VbeSession(host);
            foreach (var path in new[] { null, "", " ", "relative.tlb" })
                Assert.ThrowsException<ArgumentException>(() => session.Execute(new Request
                {
                    Command = "add_reference_file",
                    Project = project.Name,
                    Path = path
                }));
            const string selected = "{11111111-1111-1111-1111-111111111111}";
            project.References.Items.Add(new VbeSessionTests.FakeReference { GUID = selected, Major = 2, Minor = 3 });
            dynamic state = session.Execute(new Request { Command = "list_references", Project = project.Name }).Data;
            foreach (var mismatch in new[] { 0, 1, 2 })
                StringAssert.Contains(Assert.ThrowsException<InvalidOperationException>(() => session.Execute(new Request
                {
                    Command = "remove_reference",
                    Project = project.Name,
                    ExpectedReferencesVersion = state.Version,
                    Guid = mismatch == 0 ? "{22222222-2222-2222-2222-222222222222}" : selected,
                    Major = mismatch == 1 ? 4 : 2,
                    Minor = mismatch == 2 ? 4 : 3
                })).Message, "exact reference was not found");
            Assert.AreEqual(1, project.References.Items.Count);
        }
    }
}

namespace VBAi.Tests.Unit
{
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using System;
    using System.Collections;
    using System.IO;
    using System.Linq;
    using VBAi;

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

        [DataTestMethod, DataRow(0), DataRow(1), DataRow(2), DataRow(3)]
        public void ListProjectsPreservesOriginalFileNameGetterDiagnostics(int scenario)
        {
            var f = Create();
            Exception error = scenario == 1 ? new System.Runtime.InteropServices.COMException("Missing path", unchecked((int)0x800A004C)) :
                scenario == 2 ? (Exception)new DirectoryNotFoundException("Missing path") :
                scenario == 3 ? new System.Runtime.InteropServices.COMException("Unrelated failure", unchecked((int)0x80004005)) : null;
            f.Project.FileNameFailure = error;
            var response = f.Session.Execute(new Request { Command = "list_projects" });
            Assert.IsTrue(response.Ok, response.Error);
            dynamic row = ((IEnumerable)response.Data).Cast<object>().Single();
            Assert.AreEqual(f.Project.Name, (string)row.Name);
            Assert.AreEqual(f.Project.Mode, (int)row.Mode);
            Assert.AreEqual(error == null ? f.Project.FileName : null, (string)row.FileName);
            Assert.AreEqual(error?.HResult, (int?)row.FileNameErrorHResult);
            Assert.AreEqual(error?.GetType().FullName, (string)row.FileNameErrorType);
            if (error == null)
            {
                Assert.AreEqual(f.Project.FileName, (string)row.HostPath);
                Assert.IsNull((string)row.HostPathError);
            }
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
            request.Path = @"C:\DefinitelyMissing\VBAi-test.tlb";
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

namespace VBAi.Tests.Unit
{
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using System;
    using System.Collections;
    using System.Linq;
    using VBAi;

    public sealed partial class VbeSessionCoverageTests
    {
        [DataTestMethod]
        [DataRow("focus_vbe_window", typeof(ArgumentException))]
        [DataRow("show_vbe_window", typeof(ArgumentException))]
        [DataRow("window_linkage", typeof(ArgumentException))]
        [DataRow("close_vbe_window", typeof(ArgumentException))]
        [DataRow("open_object_browser", typeof(InvalidOperationException))]
        [DataRow("list_procedures", typeof(ArgumentException))]
        [DataRow("find_code", typeof(ArgumentException))]
        [DataRow("select_procedure", typeof(ArgumentException))]
        [DataRow("create_event_procedure", typeof(ArgumentException))]
        [DataRow("create_procedure", typeof(ArgumentException))]
        [DataRow("replace_procedure", typeof(ArgumentException))]
        [DataRow("remove_procedure", typeof(ArgumentException))]
        [DataRow("insert_code_file", typeof(ArgumentException))]
        [DataRow("inspect_code_file", typeof(ArgumentException))]
        [DataRow("save_host_document", typeof(ArgumentException))]
        [DataRow("save_host_document_as", typeof(ArgumentException))]
        [DataRow("read_project_signature_dialog", typeof(ArgumentException))]
        [DataRow("sign_project", typeof(ArgumentException))]
        [DataRow("component_properties", typeof(ArgumentException))]
        [DataRow("component_property_value", typeof(ArgumentException))]
        [DataRow("component_probe", typeof(ArgumentException))]
        [DataRow("set_project_property", typeof(ArgumentException))]
        [DataRow("set_component_property", typeof(ArgumentException))]
        [DataRow("set_class_instancing", typeof(ArgumentException))]
        [DataRow("rename_component", typeof(ArgumentException))]
        [DataRow("remove_component", typeof(ArgumentException))]
        [DataRow("import_component", typeof(ArgumentException))]
        [DataRow("export_component", typeof(ArgumentException))]
        [DataRow("list_reference_types", typeof(ArgumentException))]
        [DataRow("list_type_members", typeof(ArgumentException))]
        [DataRow("add_reference_guid", typeof(ArgumentException))]
        [DataRow("add_reference_file", typeof(ArgumentException))]
        [DataRow("remove_reference", typeof(ArgumentException))]
        [DataRow("read_module", typeof(ArgumentException))]
        [DataRow("create_module", typeof(ArgumentException))]
        [DataRow("create_class", typeof(ArgumentException))]
        [DataRow("run_sub", typeof(ArgumentException))]
        [DataRow("compile_project", typeof(InvalidOperationException))]
        [DataRow("open_debug_pane", typeof(ArgumentException))]
        [DataRow("add_watch", typeof(ArgumentException))]
        [DataRow("edit_watch", typeof(ArgumentException))]
        [DataRow("quick_watch", typeof(ArgumentException))]
        [DataRow("read_debug_options", typeof(InvalidOperationException))]
        [DataRow("read_vbe_options", typeof(InvalidOperationException))]
        [DataRow("set_vbe_option", typeof(InvalidOperationException))]
        [DataRow("remove_watch", typeof(ArgumentException))]
        [DataRow("debug_global", typeof(InvalidOperationException))]
        [DataRow("select_code", typeof(ArgumentException))]
        [DataRow("select_code_range", typeof(ArgumentException))]
        [DataRow("invoke_debug", typeof(ArgumentException))]
        [DataRow("form_state", typeof(ArgumentException))]
        [DataRow("form_tree", typeof(ArgumentException))]
        [DataRow("form_list_items", typeof(ArgumentException))]
        [DataRow("set_form_list_initializer", typeof(ArgumentException))]
        [DataRow("probe_append_form_list_item", typeof(ArgumentException))]
        [DataRow("add_form_list_item", typeof(ArgumentException))]
        [DataRow("remove_form_list_item", typeof(ArgumentException))]
        [DataRow("form_event_catalog", typeof(ArgumentException))]
        [DataRow("form_parent_probe", typeof(ArgumentException))]
        [DataRow("form_properties", typeof(ArgumentException))]
        [DataRow("set_form_property", typeof(ArgumentException))]
        [DataRow("set_form_picture", typeof(ArgumentException))]
        [DataRow("form_control_properties", typeof(ArgumentException))]
        [DataRow("create_form", typeof(ArgumentException))]
        [DataRow("open_form", typeof(ArgumentException))]
        [DataRow("add_form_control", typeof(ArgumentException))]
        [DataRow("add_nested_form_control", typeof(ArgumentException))]
        [DataRow("set_form_node_property", typeof(ArgumentException))]
        [DataRow("set_form_node_picture", typeof(ArgumentException))]
        [DataRow("z_order_form_control", typeof(ArgumentException))]
        [DataRow("form_property_accessors", typeof(ArgumentException))]
        [DataRow("duplicate_form_label", typeof(ArgumentException))]
        [DataRow("duplicate_form_textbox", typeof(ArgumentException))]
        [DataRow("duplicate_form_checkbox", typeof(ArgumentException))]
        [DataRow("duplicate_form_togglebutton", typeof(ArgumentException))]
        [DataRow("duplicate_form_commandbutton", typeof(ArgumentException))]
        [DataRow("duplicate_form_combobox", typeof(ArgumentException))]
        [DataRow("duplicate_empty_form_frame", typeof(ArgumentException))]
        [DataRow("frame_copy_plan", typeof(ArgumentException))]
        [DataRow("duplicate_form_frame_labels", typeof(ArgumentException))]
        [DataRow("frame_simple_copy_plan", typeof(ArgumentException))]
        [DataRow("duplicate_form_frame_simple_children", typeof(ArgumentException))]
        [DataRow("frame_profile_copy_plan", typeof(ArgumentException))]
        [DataRow("duplicate_form_frame_profiled", typeof(ArgumentException))]
        [DataRow("duplicate_form_optionbutton", typeof(ArgumentException))]
        [DataRow("remove_form_control", typeof(ArgumentException))]
        [DataRow("add_form_page", typeof(ArgumentException))]
        [DataRow("add_form_tab", typeof(ArgumentException))]
        [DataRow("remove_form_page_tab", typeof(ArgumentException))]
        [DataRow("set_form_control_geometry", typeof(ArgumentException))]
        [DataRow("rename_form_control", typeof(ArgumentException))]
        [DataRow("set_form_control_caption", typeof(ArgumentException))]
        [DataRow("set_form_control_font", typeof(ArgumentException))]
        public void PublishedCommandRejectsIncompleteArgumentsBeforeAnyHostMutation(string command, Type expected)
        {
            using (var fixture = new SigningFixture())
            {
                for (int attempt = 0; attempt < 2; attempt++)
                {
                    Exception rejection = null;
                    try { fixture.Session.Execute(new Request { Command = command, Project = fixture.Project.Name }); }
                    catch (Exception error) { rejection = error; }
                    Assert.IsNotNull(rejection, command + " accepted an incomplete request.");
                    Assert.AreEqual(expected, rejection.GetType(), command + ": " + rejection.Message);
                    Assert.IsFalse(string.IsNullOrWhiteSpace(rejection.Message), command);
                    Assert.IsFalse(rejection.Message.Contains("Unknown command"), command);
                }
                Assert.AreEqual(0, fixture.Project.VBComponents.Items.Count, command);
                Assert.AreEqual(0, fixture.Project.References.Items.Count, command);
                Assert.AreEqual(0, fixture.Scheduled, command);
                Assert.IsTrue(fixture.Project.Saved, command);
                Assert.AreEqual(2, fixture.Project.Mode, command);
            }
        }

        [TestMethod]
        public void ReadOnlyRoutesReturnTypedEmptySnapshotsAndNeverScheduleOrMutateTheHost()
        {
            using (var fixture = new SigningFixture())
            {
                foreach (var command in new[] { "status", "list_projects", "list_modules", "vbe_windows", "vbe_environment",
                    "list_addins", "code_panes", "project_properties", "project_persistence_status", "project_signature_status",
                    "list_signing_certificates", "list_references", "debug_state", "list_commands", "list_forms", "list_form_control_types" })
                    for (int attempt = 0; attempt < 2; attempt++)
                    {
                        var response = fixture.Session.Execute(new Request { Command = command, Project = fixture.Project.Name });
                        Assert.IsTrue(response.Ok, command);
                        Assert.IsNull(response.Error, command);
                        Assert.IsNotNull(response.Data, command);
                        dynamic data = response.Data;
                        if (command == "status") { Assert.IsTrue((bool)data.Connected); Assert.AreEqual("0.1.0", (string)data.Version); }
                        if (command == "list_projects") Assert.AreEqual(fixture.Project.Name, (string)((IEnumerable)response.Data).Cast<dynamic>().Single().Name);
                        if (command == "list_modules" || command == "list_commands" || command == "list_signing_certificates") Assert.AreEqual(0, ((IEnumerable)response.Data).Cast<object>().Count(), command);
                        if (command == "list_addins") Assert.AreEqual(0, (int)data.Count);
                        if (command == "vbe_environment") Assert.AreEqual("7.1", (string)data.Properties["Version"]);
                        if (command == "debug_state") Assert.AreEqual(2, (int)data.Mode);
                        if (command == "project_persistence_status" || command == "project_signature_status") Assert.AreEqual(fixture.Project.Name, (string)data.Project);
                    }
                Assert.AreEqual(0, fixture.Project.VBComponents.Items.Count);
                Assert.AreEqual(0, fixture.Project.References.Items.Count);
                Assert.AreEqual(0, fixture.Scheduled);
            }
        }

        [TestMethod]
        public void DisabledRenameAndIncompleteEditReturnExplicitFailuresRatherThanUnknownCommand()
        {
            using (var fixture = new SigningFixture())
                foreach (var command in new[] { "rename_project", "replace_lines" })
                    for (int attempt = 0; attempt < 2; attempt++)
                    {
                        var response = fixture.Session.Execute(new Request { Command = command, Project = fixture.Project.Name });
                        Assert.IsFalse(response.Ok);
                        Assert.IsNull(response.Data);
                        StringAssert.Contains(response.Error, command == "rename_project" ? "rename is disabled" : "ExpectedSha256 is required");
                        Assert.AreEqual("SigningProject", fixture.Project.Name);
                    }
        }
    }
}

namespace VBAi.Tests.Unit
{
    public sealed partial class VbeProcedureMutationTests
    {
        [Microsoft.VisualStudio.TestTools.UnitTesting.TestMethod]
        public void SessionRoutesLocalRenameThroughLiveCatalogAndVerifiedMutation()
        {
            var fixture = new Fixture("Sub Run()\r\nDim value As Long\r\nDebug.Print value\r\nEnd Sub");
            var host = new FakeVbe(); host.VBProjects.Add(fixture.Project);
            var session = new VBAi.VbeSession(host);
            var request = fixture.Request(null); request.Query = "value"; request.NewName = "amount";
            request.StartLine = 2; request.StartColumn = 5; request.ExpectedMode = 2;
            request.Command = "preview_local_rename";
            dynamic preview = session.Execute(request).Data;
            Assert.IsTrue((bool)preview.Changed); StringAssert.Contains(fixture.Module.Code, "Dim value");
            request.Command = "apply_local_rename";
            Assert.IsTrue(session.Execute(request).Ok); StringAssert.Contains(fixture.Module.Code, "Dim amount");
        }

        /// <summary>Traverse le répartiteur réel, le catalogue VBIDE et l'historique lors du renommage d'un paramètre privé.</summary>
        [Microsoft.VisualStudio.TestTools.UnitTesting.TestMethod]
        public void SessionRoutesPrivateParameterRenameAndItsManagedUndo()
        {
            const string original = "Private Sub Run(ByVal value As Long)\r\nDebug.Print value\r\nEnd Sub\r\nPublic Sub Caller()\r\nRun value:=7\r\nEnd Sub";
            var fixture = new Fixture(original);
            var host = new FakeVbe(); host.VBProjects.Add(fixture.Project);
            var session = new VBAi.VbeSession(host);
            var request = fixture.Request(null); request.Query = "value"; request.NewName = "amount";
            request.StartLine = 1; request.StartColumn = original.IndexOf("value", System.StringComparison.Ordinal) + 1; request.ExpectedMode = 2;
            request.Command = "preview_parameter_rename";
            dynamic preview = session.Execute(request).Data;
            Assert.IsTrue((bool)preview.Changed); Assert.AreEqual(original, fixture.Module.Code);
            request.Command = "apply_parameter_rename";
            Assert.IsTrue(session.Execute(request).Ok); StringAssert.Contains(fixture.Module.Code, "Run amount:=7");
            request.Command = "undo_code_edit"; request.ExpectedSha256 = Hash(fixture.Module.Code);
            Assert.IsTrue(session.Execute(request).Ok); Assert.AreEqual(original, fixture.Module.Code);
            Assert.ThrowsException<System.IO.FileNotFoundException>(() => session.Execute(new VBAi.Request
            {
                Command = "verify_vba_signature_file",
                Path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), System.Guid.NewGuid() + ".xlsm")
            }));
        }
    }

    public sealed partial class VbeDebugTests
    {
        [Microsoft.VisualStudio.TestTools.UnitTesting.TestMethod]
        public void SessionProjectPropertiesRouteChecksCapturedVersionBeforeAndDuringUiDelivery()
        {
            var previous = System.Threading.SynchronizationContext.Current;
            try
            {
                foreach (string fault in new[] { "none", "stale initially", "changed while queued" })
                {
                    var host = new VBAi.Tests.Infrastructure.EditorVbeContract();
                    var context = new NativeNavigationContext(); System.Threading.SynchronizationContext.SetSynchronizationContext(context);
                    var bar = new FakeBar { Name = "Tools" }; var control = new FakeControl { Id = 2578, Caption = "Properties" }; bar.Controls.Add(control); host.Vbe.CommandBars.Add(bar);
                    var session = new VBAi.VbeSession(host.Vbe);
                    dynamic snapshot = session.Execute(new VBAi.Request { Command = "project_properties", Project = host.Project.Name }).Data;
                    var request = new VBAi.Request { Command = "read_project_protection", Project = host.Project.Name, ExpectedMode = 2, ExpectedProjectVersion = fault == "stale initially" ? "stale" : snapshot.Version, ControlCaption = "Properties" };
                    if (fault == "stale initially") Assert.ThrowsException<System.InvalidOperationException>(() => session.Execute(request));
                    else
                    {
                        Assert.IsTrue(session.Execute(request).Ok);
                        if (fault == "changed while queued") host.Original.Name = "ChangedModule";
                        context.RunAll();
                    }
                    Assert.AreEqual(fault == "none" ? 1 : 0, control.ExecuteCount, fault);
                }
            }
            finally { System.Threading.SynchronizationContext.SetSynchronizationContext(previous); }
        }

        [Microsoft.VisualStudio.TestTools.UnitTesting.TestMethod]
        public void SessionRoutesProcedureQueueAndStatusThroughBoundDebugger()
        {
            var previous = System.Threading.SynchronizationContext.Current;
            var context = new NativeNavigationContext(); System.Threading.SynchronizationContext.SetSynchronizationContext(context);
            try
            {
                var fixture = Create(2); var request = ProcedureRequest(fixture);
                fixture.Service.ShowProcedureImmediate = () => { };
                fixture.Service.ExecuteProcedureCall = command => System.Threading.Tasks.Task.FromResult<object>("delivered");
                var session = new VBAi.VbeSession(fixture.Vbe);
                typeof(VBAi.VbeSession).GetField("debugger", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic).SetValue(session, fixture.Service);
                request.Command = "run_procedure"; dynamic queued = session.Execute(request).Data;
                context.RunAll();
                request.Command = "procedure_run_status"; request.Query = queued.Query;
                dynamic status = session.Execute(request).Data;
                Assert.AreEqual("Delivered", (string)status.State); Assert.IsFalse((bool)status.Pending);
            }
            finally { System.Threading.SynchronizationContext.SetSynchronizationContext(previous); }
        }
    }
}
