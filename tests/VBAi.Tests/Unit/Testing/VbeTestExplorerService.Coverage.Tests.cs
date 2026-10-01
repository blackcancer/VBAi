using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Unit
{
    public sealed partial class VbeTestExplorerServiceTests
    {
        [STATestMethod]
        public void CoverageRejectsUnsupportedDocumentFormatBeforeCreatingOrInvokingACopy()
        {
            using (var fixture = new Fixture())
            using (var coverage = new CoverageFixture(fixture))
            {
                fixture.Project.FileName = @"C:\Temp\Fixture.xlsx";
                var catalog = fixture.Catalog();
                StringAssert.Contains(fixture.Service.CoverageUnavailableReason(catalog), "document format");
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.RunCoverageAsync(catalog, catalog.Tests.ToArray(), null, CancellationToken.None));
                Assert.AreEqual(0, coverage.Events.Count);
                Assert.AreEqual(0, fixture.Host.Invocations);
            }
        }

        [STATestMethod]
        public void CoverageInstrumentsOnlyOwnedCopyMapsOriginalIdsAndCollectsBeforeClosing()
        {
            using (var fixture = new Fixture())
            using (var coverage = new CoverageFixture(fixture))
            {
                var catalog = fixture.Catalog();
                string original = fixture.Project.VBComponents.Single(component => component.Name == "Production").CodeModule.Source;
                var run = Pump(fixture.Service.RunCoverageAsync(catalog, catalog.Tests.ToArray(), null, CancellationToken.None));
                Assert.AreEqual(original, fixture.Project.VBComponents.Single(component => component.Name == "Production").CodeModule.Source);
                Assert.AreEqual(0, fixture.Project.VBComponents.Single(component => component.Name == "Production").CodeModule.Insertions);
                Assert.IsTrue(coverage.Clone.VBComponents.Single(component => component.Name == "Production").CodeModule.Source.Contains(
                    VbaCoverageInstrumentation.ModuleName + "." + VbaCoverageInstrumentation.HitsVariable));
                CollectionAssert.AreEqual(new[] { "Compile", "Reset", "Test:Alpha", "Test:Beta", "Snapshot", "Close" }, coverage.Events);
                CollectionAssert.AreEquivalent(catalog.Tests.Select(test => test.Id).ToArray(), run.Results.Select(result => result.Test.Id).ToArray());
                Assert.IsTrue(run.Coverage.Available);
                Assert.IsTrue(run.Coverage.Complete);
                Assert.AreEqual(50d, run.Coverage.Percent.Value);
                Assert.AreEqual(catalog.Project.Revision, run.Coverage.Revision);
                Assert.AreEqual(catalog.Project.Selector, run.Coverage.Original);
            }
        }

        [STATestMethod]
        public void CoverageRejectsOriginalIdentityAndPathWithoutCallingTheirCloseOrWritingCode()
        {
            foreach (bool originalIdentity in new[] { false, true })
            using (var fixture = new Fixture())
            using (var coverage = new CoverageFixture(fixture))
            {
                fixture.Service.CreateCoverageClone = (project, path, folder) => new VbaTestCoverageClone
                {
                    Project = originalIdentity ? fixture.Project : coverage.MakeCopy(folder),
                    Path = fixture.Project.FileName,
                    Close = () => coverage.Events.Add("UnsafeClose")
                };
                var catalog = fixture.Catalog();
                var run = Pump(fixture.Service.RunCoverageAsync(catalog, catalog.Tests.ToArray(), null, CancellationToken.None));
                Assert.IsFalse(run.Coverage.Available);
                Assert.IsFalse(coverage.Events.Contains("UnsafeClose"));
                Assert.IsFalse(coverage.Events.Contains("Reset"));
                Assert.IsTrue(fixture.Project.VBComponents.All(component => component.CodeModule.Insertions == 0));
            }
        }

        [STATestMethod]
        public void CoverageRejectsChangedCopiedTestsAndReferencesBeforeAnyInstrumentationOrNativeCall()
        {
            foreach (bool reference in new[] { false, true })
            using (var fixture = new Fixture())
            using (var coverage = new CoverageFixture(fixture))
            {
                coverage.AfterCopy = copy =>
                {
                    if (reference) copy.References[0].Minor++;
                    else copy.VBComponents.Single(component => component.Name == "TestsOne").CodeModule.Source += "\n' wrong test revision";
                };
                var catalog = fixture.Catalog();
                var run = Pump(fixture.Service.RunCoverageAsync(catalog, catalog.Tests.ToArray(), null, CancellationToken.None));
                Assert.IsFalse(run.Coverage.Available);
                Assert.IsFalse(coverage.Events.Contains("Reset"));
                Assert.IsTrue(coverage.Clone.VBComponents.All(component => component.CodeModule.Insertions == 0));
                Assert.AreEqual(1, coverage.Events.Count(item => item == "Close"));
            }
        }

        [STATestMethod]
        public void CoverageStopsFurtherCopyWritesWhenOriginalChangesDuringAnInstrumentedModuleEdit()
        {
            using (var fixture = new Fixture())
            using (var coverage = new CoverageFixture(fixture))
            {
                coverage.AfterCopy = copy => copy.VBComponents.Single(component => component.Name == "Production").CodeModule.OnInsert =
                    () => fixture.Project.VBComponents.Single(component => component.Name == "TestsOne").CodeModule.Source += "\n' changed during clone edit";
                var catalog = fixture.Catalog();
                var run = Pump(fixture.Service.RunCoverageAsync(catalog, catalog.Tests.ToArray(), null, CancellationToken.None));
                Assert.IsFalse(run.Coverage.Available);
                Assert.IsFalse(coverage.Clone.VBComponents.Any(component => component.Name == VbaCoverageInstrumentation.ModuleName));
                Assert.AreEqual(0, coverage.Clone.VBComponents.Single(component => component.Name == VbaTestRuntimeSource.ModuleName).CodeModule.Insertions);
                Assert.IsFalse(coverage.Events.Contains("Reset"));
            }
        }

        [STATestMethod]
        public void CoverageRequiresVerifiedResetAndNeverRunsTestsAfterInvalidResetReturn()
        {
            using (var fixture = new Fixture())
            using (var coverage = new CoverageFixture(fixture))
            {
                coverage.ResetResult = false;
                var catalog = fixture.Catalog();
                var run = Pump(fixture.Service.RunCoverageAsync(catalog, catalog.Tests.ToArray(), null, CancellationToken.None));
                Assert.IsFalse(run.Coverage.Available);
                Assert.IsTrue(coverage.Events.Contains("Reset"));
                Assert.IsFalse(coverage.Events.Any(item => item.StartsWith("Test:", StringComparison.Ordinal)));
                Assert.IsFalse(coverage.Events.Contains("Snapshot"));
            }
        }

        [STATestMethod]
        public void CoverageDoesNotDispatchTestsWhenCloneCompilationIsRefused()
        {
            using (var fixture = new Fixture())
            using (var coverage = new CoverageFixture(fixture))
            {
                fixture.Service.CompileCoverageProject = project => { coverage.Events.Add("Compile"); throw new InvalidOperationException("Clone compilation refused"); };
                var catalog = fixture.Catalog();
                var run = Pump(fixture.Service.RunCoverageAsync(catalog, catalog.Tests.ToArray(), null, CancellationToken.None));
                Assert.IsFalse(run.Coverage.Available);
                Assert.IsFalse(coverage.Events.Contains("Reset"));
                Assert.IsTrue(run.Results.All(result => result.Outcome == VbaTestOutcome.Blocked));
                StringAssert.Contains(run.Error, "compilation refused");
            }
        }

        [STATestMethod]
        public void UncertainCoveredTestNeverCollectsAnotherSnapshotOrClosesTheUnreconciledCopy()
        {
            using (var fixture = new Fixture())
            using (var coverage = new CoverageFixture(fixture))
            {
                coverage.UnknownTest = true;
                var catalog = fixture.Catalog();
                var run = Pump(fixture.Service.RunCoverageAsync(catalog, catalog.Tests.ToArray(), null, CancellationToken.None));
                Assert.IsTrue(run.OutcomeUnknown);
                Assert.IsFalse(run.Coverage.Available);
                Assert.IsNull(run.Coverage.Hit);
                Assert.IsNull(run.Coverage.Percent);
                CollectionAssert.AreEqual(new[] { "Compile", "Reset", "Test:Alpha" }, coverage.Events);
                Assert.IsTrue(fixture.Vbe.VBProjects.Contains(coverage.Clone), "Keep the uncertain copy available for inspection.");
            }
        }

        [STATestMethod]
        public void StoppedCoverageRetainsVerifiedPartialHitsAndDoesNotRunQueuedTest()
        {
            using (var fixture = new Fixture())
            using (var coverage = new CoverageFixture(fixture))
            using (var cancellation = new CancellationTokenSource())
            {
                coverage.AfterTest = () => cancellation.Cancel();
                var catalog = fixture.Catalog();
                var run = Pump(fixture.Service.RunCoverageAsync(catalog, catalog.Tests.ToArray(), null, cancellation.Token));
                Assert.IsTrue(run.Coverage.Available);
                Assert.IsFalse(run.Coverage.Complete);
                Assert.AreEqual(1, run.Coverage.Hit);
                Assert.IsNull(run.Coverage.Percent);
                Assert.AreEqual(VbaTestOutcome.Passed, run.Results[0].Outcome);
                Assert.AreEqual(VbaTestOutcome.Cancelled, run.Results[1].Outcome);
                CollectionAssert.AreEqual(new[] { "Compile", "Reset", "Test:Alpha", "Snapshot", "Close" }, coverage.Events);
            }
        }

        [STATestMethod]
        public void FailedCopyCloseDoesNotInventSuccessfulCleanupOrCompleteCoverage()
        {
            using (var fixture = new Fixture())
            using (var coverage = new CoverageFixture(fixture))
            {
                coverage.CloseFails = true;
                var catalog = fixture.Catalog();
                var run = Pump(fixture.Service.RunCoverageAsync(catalog, catalog.Tests.ToArray(), null, CancellationToken.None));
                Assert.IsTrue(run.Coverage.Available);
                Assert.IsFalse(run.Coverage.Complete);
                Assert.IsNull(run.Coverage.Percent);
                StringAssert.Contains(run.Error, "could not be closed");
                Assert.IsTrue(fixture.Vbe.VBProjects.Contains(coverage.Clone));
            }
        }

        [STATestMethod]
        public void PermissionRevokedDuringTheLastCoverageSnapshotCannotPublishCompleteCoverage()
        {
            using (var fixture = new Fixture())
            using (var coverage = new CoverageFixture(fixture))
            {
                bool permitted = true;
                coverage.BeforeSnapshotReturn = () => permitted = false;
                var catalog = fixture.Catalog();
                dynamic start = fixture.Service.StartRun(fixture.Project.FileName, catalog.Project.Revision,
                    catalog.Tests.Select(test => test.Id).ToArray(),
                    () => { if (!permitted) throw new InvalidOperationException("Execution permission was revoked."); }, true);
                string id = start.Query;
                PumpMessagesUntil(() => !((bool)((dynamic)fixture.Service.RunStatus(fixture.Project.FileName, id, "compact")).Pending));
                dynamic status = fixture.Service.RunStatus(fixture.Project.FileName, id, "compact");
                var report = (Dictionary<string, object>)status.Report;
                var measured = (Dictionary<string, object>)report["coverage"];
                Assert.IsTrue((bool)report["uncertain"]);
                Assert.IsFalse((bool)measured["available"]);
                Assert.IsFalse((bool)measured["complete"]);
                Assert.IsNull(measured["percent"]);
                Assert.AreEqual(1, coverage.Events.Count(item => item == "Snapshot"));
                Assert.IsFalse(coverage.Events.Contains("Close"));
                Assert.IsTrue(fixture.Vbe.VBProjects.Contains(coverage.Clone));
            }
        }

        [STATestMethod]
        public void SourceChangesDuringTheLastCoverageSnapshotInvalidateItsMeasurementAndRetainTheCopy()
        {
            foreach (bool originalChanged in new[] { true, false })
            using (var fixture = new Fixture())
            using (var coverage = new CoverageFixture(fixture))
            {
                coverage.BeforeSnapshotReturn = () =>
                    (originalChanged ? fixture.Project : coverage.Clone).VBComponents[0].CodeModule.Source += "\n' changed during snapshot";
                var catalog = fixture.Catalog();
                var run = Pump(fixture.Service.RunCoverageAsync(catalog, catalog.Tests.ToArray(), null, CancellationToken.None));
                Assert.IsTrue(run.OutcomeUnknown);
                Assert.IsFalse(run.Coverage.Available);
                Assert.IsFalse(run.Coverage.Complete);
                Assert.IsNull(run.Coverage.Percent);
                Assert.IsFalse(coverage.Events.Contains("Close"));
                Assert.AreEqual(1, coverage.Events.Count(item => item == "Snapshot"));
                Assert.IsTrue(fixture.Vbe.VBProjects.Contains(coverage.Clone));
            }
        }

        private sealed class CoverageFixture : VbeDebug.IProcedureValuesHost, IDisposable
        {
            private readonly Fixture original;
            private readonly string folder = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "VBAi-coverage-test-" + Guid.NewGuid().ToString("N"));
            internal readonly List<string> Events = new List<string>();
            internal FakeProject Clone;
            internal Action<FakeProject> AfterCopy;
            internal Action AfterTest;
            internal Action BeforeSnapshotReturn;
            internal bool UnknownTest, CloseFails;
            internal object ResetResult = true;
            internal CoverageFixture(Fixture original)
            {
                this.original = original;
                original.Project.VBComponents.Add(new FakeComponent { Name = "Production", Type = 1,
                    CodeModule = new FakeCode { Source = "Public Sub One()\nEnd Sub\nPublic Sub Two()\nEnd Sub" } });
                original.InstallFixtureSupport();
                original.Service.Host = this;
                original.Service.CompileCoverageProject = project => { Assert.AreSame(Clone, project); Events.Add("Compile"); };
                original.Service.CoverageRoot = () => folder;
                original.Service.CreateCoverageClone = (project, path, destination) =>
                {
                    var copy = MakeCopy(destination);
                    AfterCopy?.Invoke(copy);
                    return new VbaTestCoverageClone { Project = copy, Path = copy.FileName,
                        Close = () => { Events.Add("Close"); if (CloseFails) throw new InvalidOperationException("Simulated close failure"); original.Vbe.VBProjects.Remove(copy); } };
                };
            }
            internal FakeProject MakeCopy(string destination)
            {
                Clone = new FakeProject { Name = original.Project.Name, FileName = System.IO.Path.Combine(destination, "coverage.xlsm") };
                foreach (var component in original.Project.VBComponents)
                    Clone.VBComponents.Add(new FakeComponent { Name = component.Name, Type = component.Type,
                        CodeModule = new FakeCode { Source = component.CodeModule.Source } });
                foreach (var reference in original.Project.References)
                    Clone.References.Add(new FakeReference { Name = reference.Name, Guid = reference.Guid,
                        Major = reference.Major, Minor = reference.Minor, IsBroken = reference.IsBroken });
                original.Vbe.VBProjects.Add(Clone);
                return Clone;
            }
            public object ResolveTarget(object project, string expectedHostPath)
            {
                Assert.AreEqual(((FakeProject)project).FileName, expectedHostPath);
                return project;
            }
            public object Invoke(object target, string module, string procedure, object[] arguments)
            {
                Assert.AreSame(Clone, target, "Coverage may invoke only the owned copy.");
                if (module == VbaCoverageInstrumentation.ModuleName)
                {
                    if (procedure == VbaCoverageInstrumentation.ResetProcedure) { Events.Add("Reset"); return ResetResult; }
                    Assert.AreEqual(VbaCoverageInstrumentation.SnapshotProcedure, procedure);
                    Events.Add("Snapshot");
                    var hits = Array.CreateInstance(typeof(bool), new[] { 2 }, new[] { 1 });
                    hits.SetValue(true, 1); hits.SetValue(false, 2);
                    BeforeSnapshotReturn?.Invoke();
                    return hits;
                }
                Assert.AreEqual(VbaTestRuntimeSource.ModuleName, module);
                Assert.AreEqual(VbaTestRuntimeSource.DispatcherProcedure, procedure);
                Events.Add("Test:" + arguments[1]);
                if (UnknownTest) throw new InvalidOperationException("No verified covered test completion");
                AfterTest?.Invoke();
                return new object[] { "Passed", "", "0" };
            }
            public void Dispose()
            {
                if (Clone != null) original.Vbe.VBProjects.Remove(Clone);
                if (Directory.Exists(folder)) Directory.Delete(folder, true);
            }
        }
    }
}
