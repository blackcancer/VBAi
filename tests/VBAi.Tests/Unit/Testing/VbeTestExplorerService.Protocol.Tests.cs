using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Linq;
using System.Runtime.ExceptionServices;
using System.Threading.Tasks;

namespace VBAi.Tests.Unit
{
    public sealed partial class VbeTestExplorerServiceTests
    {
        private const string RevisionRefusal = "ExpectedProjectVersion must match the discovered test project revision.";

        /// <summary>Every guarded route refuses missing or stale source versions without emitting a caught validation exception.</summary>
        [STATestMethod]
        [DataRow("run_vba_tests", null), DataRow("run_vba_tests", ""), DataRow("run_vba_tests", "stale")]
        [DataRow("install_vba_test_support", null), DataRow("install_vba_test_support", ""), DataRow("install_vba_test_support", "stale")]
        [DataRow("navigate_vba_test", null), DataRow("navigate_vba_test", ""), DataRow("navigate_vba_test", "stale")]
        public void BridgeRefusesInvalidTestRevisionsWithoutDispatchOrFirstChanceValidation(string command, string revision)
        {
            using (var fixture = new Fixture())
            {
                fixture.InstallFixtureSupport();
                var catalog = fixture.Catalog();
                string source = fixture.Project.VBComponents[0].CodeModule.Source;
                int nativeCalls = fixture.Host.Invocations, approvals = 0, validationExceptions = 0;
                EventHandler<FirstChanceExceptionEventArgs> observer = (_, args) =>
                {
                    if (args.Exception is InvalidOperationException && args.Exception.Message == RevisionRefusal)
                        validationExceptions++;
                };
                AppDomain.CurrentDomain.FirstChanceException += observer;
                Response reply;
                try
                {
                    reply = fixture.Service.CommandResponse(new Request
                    {
                        Command = command, Project = fixture.Project.FileName,
                        ExpectedProjectVersion = revision, ExpectedMode = 2,
                        Items = catalog.Tests.Select(test => test.Id).ToArray(), Text = "must not install"
                    }, () => approvals++);
                }
                finally { AppDomain.CurrentDomain.FirstChanceException -= observer; }
                Assert.IsFalse(reply.Ok);
                Assert.AreEqual(RevisionRefusal, reply.Error);
                Assert.IsNull(reply.Data, "A refusal must not be nested inside a successful payload.");
                Assert.AreEqual(0, validationExceptions);
                Assert.AreEqual(nativeCalls, fixture.Host.Invocations);
                Assert.AreEqual(0, approvals);
                Assert.AreEqual(source, fixture.Project.VBComponents[0].CodeModule.Source);
                Assert.AreEqual(catalog.Project.Revision, fixture.Catalog().Project.Revision);
            }
        }

        /// <summary>The actual session route exposes the same failure after source changes, retaining historical test state.</summary>
        [STATestMethod]
        [DataRow("run_vba_tests"), DataRow("install_vba_test_support"), DataRow("navigate_vba_test")]
        public void SessionDoesNotWrapAStaleTestRefusalAsSuccess(string command)
        {
            using (var fixture = new Fixture())
            {
                var session = new VbeSession(fixture.Vbe);
                fixture.InstallFixtureSupport();
                var catalog = fixture.Catalog();
                fixture.Project.VBComponents[0].CodeModule.Source += "\r\n' changed after discovery\r\n";
                string changedSource = fixture.Project.VBComponents[0].CodeModule.Source;
                session.TestExplorer = fixture.Service;
                var reply = session.Execute(new Request
                {
                    Command = command, Project = fixture.Project.FileName,
                    ExpectedProjectVersion = catalog.Project.Revision, ExpectedMode = 2,
                    Items = catalog.Tests.Select(test => test.Id).ToArray(), Text = "must not install"
                });
                Assert.IsFalse(reply.Ok);
                Assert.AreEqual(RevisionRefusal, reply.Error);
                Assert.IsNull(reply.Data);
                Assert.AreEqual(0, fixture.Host.Invocations);
                Assert.AreEqual(changedSource, fixture.Project.VBComponents[0].CodeModule.Source);
            }
        }

        /// <summary>Read-only replies retain the existing payload and do not acquire a new version requirement.</summary>
        [STATestMethod]
        [DataRow("discover_vba_tests"), DataRow("preview_vba_test_support"), DataRow("vba_test_coverage")]
        public void BridgePreservesReadOnlyTestPayloadsWithoutExpectedVersion(string command)
        {
            using (var fixture = new Fixture())
            {
                var reply = fixture.Service.CommandResponse(new Request { Command = command, Project = fixture.Project.FileName });
                Assert.IsTrue(reply.Ok);
                Assert.IsNull(reply.Error);
                Assert.IsNotNull(reply.Data);
                Assert.IsFalse(reply.Data is Response);
                Assert.AreEqual(0, fixture.Host.Invocations);
            }
        }

        /// <summary>A valid revision still reaches the original permission callback and never conceals its refusal.</summary>
        [STATestMethod]
        public void BridgePreservesTheNativeRunAuthorizationGuard()
        {
            using (var fixture = new Fixture())
            {
                fixture.InstallFixtureSupport();
                var catalog = fixture.Catalog();
                int checks = 0;
                var request = new Request
                {
                    Command = "run_vba_tests", Project = fixture.Project.FileName,
                    ExpectedProjectVersion = catalog.Project.Revision, ExpectedMode = 2,
                    Items = catalog.Tests.Select(test => test.Id).ToArray()
                };
                var error = Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.CommandResponse(
                    request, () => { checks++; throw new InvalidOperationException("Permission withdrawn"); }));
                Assert.AreEqual("Permission withdrawn", error.Message);
                Assert.AreEqual(1, checks);
                Assert.AreEqual(0, fixture.Host.Invocations);
            }
        }

        /// <summary>Owner-thread and lifetime faults remain programmer errors rather than being converted to revision refusals.</summary>
        [STATestMethod]
        public void BridgeDoesNotMaskWrongOwnerOrDisposedService()
        {
            using (var fixture = new Fixture())
            {
                var request = new Request { Command = "discover_vba_tests", Project = fixture.Project.FileName };
                var error = Task.Run(() =>
                {
                    try { fixture.Service.CommandResponse(request); return null; }
                    catch (Exception failure) { return failure; }
                }).GetAwaiter().GetResult();
                Assert.IsInstanceOfType(error, typeof(InvalidOperationException));
                StringAssert.Contains(error.Message, "owning VBE thread");
                fixture.Service.Dispose();
                Assert.ThrowsException<ObjectDisposedException>(() => fixture.Service.CommandResponse(request));
                Assert.AreEqual(0, fixture.Host.Invocations);
            }
        }
    }
}
