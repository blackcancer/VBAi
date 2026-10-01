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

        [STATestMethod]
        public void CoverageAvailabilityRejectsUnsupportedHostSyntaxModeAndUnsavedWordBeforeCopying()
        {
            using (var fixture = new Fixture())
            {
                fixture.Service.IsExecutionHost = () => false;
                StringAssert.Contains(fixture.Service.CoverageUnavailableReason(fixture.Catalog()), "Excel");
                fixture.Service.IsExecutionHost = () => true;
                fixture.Project.VBComponents.Add(new FakeComponent { Name = "Unsupported", Type = 1, CodeModule = new FakeCode { Source = "#If VBA7 Then\nPublic Sub One()\nEnd Sub\n#End If" } });
                Assert.IsNotNull(fixture.Service.CoverageUnavailableReason(fixture.Catalog()));
                fixture.Project.VBComponents.RemoveAt(1); fixture.Project.Mode = 1;
                StringAssert.Contains(fixture.Service.CoverageUnavailableReason(fixture.Catalog()), "design mode");
            }
            using (var fixture = new Fixture())
            {
                fixture.Project.FileName = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "Unsaved.docm");
                var app = new VbaTestWordValuesHostTests.Application();
                var document = new VbaTestWordValuesHostTests.Document { Application = app, FullName = fixture.Project.FileName, VBProject = fixture.Project, Saved = false };
                app.Documents.Add(document);
                fixture.Service.Host = new VbaTestWordValuesHost { ReadProcessName = () => "WINWORD", ReadProcessId = () => 123,
                    ReadActiveApplication = _ => app, ReadWindowOwner = _ => 123, SameIdentity = ReferenceEquals,
                    ReadDocumentItem = (documents, index) => ((VbaTestWordValuesHostTests.Documents)documents)[index - 1] };
                StringAssert.Contains(fixture.Service.CoverageUnavailableReason(fixture.Catalog()), "Word document changes");
                document.Saved = true; Assert.IsNull(fixture.Service.CoverageUnavailableReason(fixture.Catalog()));
                fixture.Service.Host = new VbaTestPowerPointValuesHost(); fixture.Project.FileName = "C:\\Temp\\Fixture.xlsx";
                StringAssert.Contains(fixture.Service.CoverageUnavailableReason(fixture.Catalog()), "document format");
            }
        }

        [STATestMethod]
        public void CoverageRejectsMissingEscapedOrMismatchedCloneBeforeWritingAndRetainsUnownedCopies()
        {
            foreach (string fault in new[] { "null", "project", "path", "outside", "projectPath", "moduleCount", "moduleName", "moduleType", "compile", "resetType" })
            using (var fixture = new Fixture())
            using (var coverage = new CoverageFixture(fixture))
            {
                fixture.Service.CreateCoverageClone = (project, path, folder) => {
                    if (fault == "null") return null;
                    var copy = coverage.MakeCopy(folder);
                    if (fault == "moduleCount") copy.VBComponents.RemoveAt(0);
                    if (fault == "moduleName") copy.VBComponents[0].Name = "Renamed";
                    if (fault == "moduleType") copy.VBComponents[0].Type = 2;
                    string target = copy.FileName;
                    if (fault == "projectPath") copy.FileName = System.IO.Path.Combine(folder, "Other.xlsm");
                    if (fault == "outside") target = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "Unowned.xlsm");
                    return new VbaTestCoverageClone { Project = fault == "project" ? null : copy,
                        Path = fault == "path" ? null : target, Close = () => coverage.Events.Add("Close") };
                };
                if (fault == "compile") fixture.Service.CompileCoverageProject = null;
                if (fault == "resetType") coverage.ResetResult = new object();
                var catalog = fixture.Catalog();
                var run = Pump(fixture.Service.RunCoverageAsync(catalog, catalog.Tests.ToArray(), null, CancellationToken.None));
                Assert.IsFalse(run.Coverage.Available, fault); Assert.IsFalse(run.OutcomeUnknown, fault);
                Assert.IsTrue(run.Results.All(result => result.Outcome == VbaTestOutcome.Blocked), fault);
                Assert.IsFalse(coverage.Events.Any(item => item.StartsWith("Test:", StringComparison.Ordinal)), fault);
                Assert.AreEqual(new[] { "moduleCount", "moduleName", "moduleType", "compile", "resetType" }.Contains(fault), coverage.Events.Contains("Close"), fault);
            }
        }

        [STATestMethod]
        public void CloneModuleBoundariesRejectMissingCollidingAndUnverifiedSourceAndPreserveInlineEdits()
        {
            var write = typeof(VbeTestExplorerService).GetMethod("WriteCloneModule", System.Reflection.BindingFlags.Static | System.Reflection.BindingFlags.NonPublic);
            var instrument = typeof(VbeTestExplorerService).GetMethod("InstrumentCloneModule", System.Reflection.BindingFlags.Static | System.Reflection.BindingFlags.NonPublic);
            var project = new FakeProject();
            Assert.IsInstanceOfType(Assert.ThrowsException<System.Reflection.TargetInvocationException>(() => write.Invoke(null, new object[] { project, "Missing", "source", false })).InnerException, typeof(InvalidOperationException));
            write.Invoke(null, new object[] { project, "Runtime", "first", true });
            Assert.AreEqual("first", project.VBComponents.Single().CodeModule.Source);
            Assert.IsInstanceOfType(Assert.ThrowsException<System.Reflection.TargetInvocationException>(() => write.Invoke(null, new object[] { project, "Runtime", "second", true })).InnerException, typeof(InvalidOperationException));
            write.Invoke(null, new object[] { project, "Runtime", "second", false });
            Assert.AreEqual("second", project.VBComponents.Single().CodeModule.Source);
            var module = new VbaCoverageModule { Name = "Missing", InstrumentedSource = "" };
            Assert.IsInstanceOfType(Assert.ThrowsException<System.Reflection.TargetInvocationException>(() => instrument.Invoke(null, new object[] { project, module, (Action)(() => { }) })).InnerException, typeof(InvalidOperationException));
            module.Name = "Runtime"; module.InstrumentedSource = "probe second";
            module.Edits.Add(new VbaCoverageEdit { OriginalLine = 1, OriginalColumn = 1, Text = "probe ", IsWholeLine = false });
            int guards = 0; instrument.Invoke(null, new object[] { project, module, (Action)(() => guards++) });
            Assert.AreEqual("probe second", project.VBComponents.Single().CodeModule.Source); Assert.AreEqual(1, guards);
            module.Edits.Clear(); module.InstrumentedSource = "different";
            Assert.IsInstanceOfType(Assert.ThrowsException<System.Reflection.TargetInvocationException>(() => instrument.Invoke(null, new object[] { project, module, (Action)(() => { }) })).InnerException, typeof(InvalidOperationException));
            project.VBComponents.Single().CodeModule.Source = ""; module.InstrumentedSource = "";
            instrument.Invoke(null, new object[] { project, module, (Action)(() => { }) });
        }

        [STATestMethod]
        public void UnsupportedCoveragePlanCannotAdvertiseSupportEvenWithoutAHostFailure()
        {
            var plan = new VbaCoveragePlan { DenominatorKnown = false };
            dynamic preview = VbeTestExplorerService.CoveragePreviewPage("Disposable", "r1", plan, null);
            Assert.IsFalse((bool)preview.Supported);
            Assert.IsFalse((bool)preview.Available);
            Assert.IsFalse((bool)preview.StatementCoverageAvailable);
        }

        [STATestMethod]
        public void CoveredTestsAcrossModulesKeepTheirOriginalIdentityAndPublishEveryProgressResult()
        {
            using (var fixture = new Fixture())
            {
                fixture.Project.VBComponents.Add(new FakeComponent { Name = "TestsTwo", Type = 1,
                    CodeModule = new FakeCode { Source = "'@TestModule\n'@TestMethod\nPublic Sub OtherModuleTest()\nEnd Sub" } });
                using (var coverage = new CoverageFixture(fixture))
                {
                    var catalog = fixture.Catalog(); var progress = new List<VbaTestResult>();
                    var run = Pump(fixture.Service.RunCoverageAsync(catalog, catalog.Tests.ToArray(), progress.Add, CancellationToken.None));
                    Assert.AreEqual(3, run.Results.Count);
                    CollectionAssert.AreEqual(run.Results.Select(result => result.Test.Id).ToArray(), progress.Select(result => result.Test.Id).ToArray());
                    Assert.AreSame(catalog.Tests.Single(test => test.Module == "TestsTwo"), run.Results.Single(result => result.Test.Module == "TestsTwo").Test);
                    Assert.IsTrue(run.Coverage.Complete);
                    dynamic report = fixture.Service.Command(new Request { Command = "vba_test_coverage", Project = fixture.Project.FileName, Query = run.Id });
                    Assert.IsTrue((bool)report.Available);
                }
            }
        }

        [STATestMethod]
        public void NativeCompilerBoundaryIsObservedOnOwnerBeforeResetAndCancellationRefusesDispatch()
        {
            foreach (bool cancelDuringCompile in new[] { false, true })
            using (var fixture = new Fixture())
            using (var coverage = new CoverageFixture(fixture))
            using (var cancellation = new CancellationTokenSource())
            {
                var command = new VbeTestExplorerServiceCompilationTests.CompileControl();
                fixture.Vbe.CommandBars.Control = command;
                coverage.AfterCopy = copy => fixture.Vbe.ActiveVBProject = copy;
                command.OnExecute = () => { coverage.Events.Add("Compile"); command.State = false; if (cancelDuringCompile) cancellation.Cancel(); };
                fixture.Service.CompileCoverageProject = (Action<object>)typeof(VbeTestExplorerService)
                    .GetField("defaultCompileCoverageDelegate", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic).GetValue(fixture.Service);
                var catalog = fixture.Catalog();
                var progress = new List<VbaTestResult>();
                var run = Pump(fixture.Service.RunCoverageAsync(catalog, catalog.Tests.ToArray(), progress.Add, cancellation.Token));
                Assert.AreEqual(1, command.Executions, "Neither observing compilation nor cancellation may retry the native command. " + run.Error);
                Assert.AreEqual(!cancelDuringCompile, run.Coverage.Available);
                Assert.AreEqual(!cancelDuringCompile, coverage.Events.Contains("Reset"));
                Assert.AreEqual(catalog.Tests.Count(), progress.Count);
                if (cancelDuringCompile) Assert.IsTrue(run.Results.All(result => result.Outcome == VbaTestOutcome.Cancelled));
                else Assert.IsTrue(run.Results.All(result => result.Outcome == VbaTestOutcome.Passed));
                Assert.IsTrue(coverage.Events.Contains("Close"));
            }
        }

        [STATestMethod]
        public void OwnerDispatcherDisappearingDuringCompilationObserverSetupProducesKnownRefusal()
        {
            using (var fixture = new Fixture())
            using (var coverage = new CoverageFixture(fixture))
            {
                var command = new VbeTestExplorerServiceCompilationTests.CompileControl { State = false };
                fixture.Vbe.CommandBars.Control = command;
                coverage.AfterCopy = copy => fixture.Vbe.ActiveVBProject = copy;
                fixture.Service.CompileCoverageProject = (Action<object>)typeof(VbeTestExplorerService)
                    .GetField("defaultCompileCoverageDelegate", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic).GetValue(fixture.Service);
                // The observer's setup boundary loses its external UI handle after its initial owner check.
                // Its already faulted completion must still settle through the independent run dispatcher.
                fixture.Service.CoverageCompilationClock = () => { fixture.Dispatcher.Dispose(); return 0; };
                var catalog = fixture.Catalog();
                var run = Pump(fixture.Service.RunCoverageAsync(catalog, catalog.Tests.ToArray(), null, CancellationToken.None));
                Assert.IsFalse(run.Coverage.Available); Assert.IsFalse(run.OutcomeUnknown);
                Assert.IsFalse(coverage.Events.Contains("Reset"));
                Assert.AreEqual(0, command.Executions);
                Assert.IsTrue(coverage.Events.Contains("Close"));
            }
        }

        [STATestMethod]
        public void CopyCreationCancellationPublishesCancelledPreparationForEverySelectedTest()
        {
            using (var fixture = new Fixture())
            using (var coverage = new CoverageFixture(fixture))
            using (var cancellation = new CancellationTokenSource())
            {
                coverage.AfterCopy = _ => cancellation.Cancel();
                var catalog = fixture.Catalog(); var progress = new List<VbaTestResult>();
                var run = Pump(fixture.Service.RunCoverageAsync(catalog, catalog.Tests.ToArray(), progress.Add, cancellation.Token));
                Assert.AreEqual(2, progress.Count);
                Assert.IsTrue(progress.All(result => result.Outcome == VbaTestOutcome.Cancelled && result.Phase == "Preparation"));
                Assert.IsFalse(run.Coverage.Available); Assert.IsFalse(run.OutcomeUnknown);
                Assert.IsFalse(coverage.Events.Contains("Reset"));
                Assert.IsTrue(coverage.Events.Contains("Close"));
            }
        }

        [STATestMethod]
        public void CloseFailuresAppendPreparationErrorAndPreserveKnownOrUncertainCompletion()
        {
            foreach (bool uncertain in new[] { false, true })
            using (var fixture = new Fixture())
            using (var coverage = new CoverageFixture(fixture))
            {
                fixture.Service.CompileCoverageProject = _ => { throw new InvalidOperationException("Preparation refused"); };
                coverage.CloseError = new VbaTestInvocationException("Copy close lost completion", uncertain);
                var catalog = fixture.Catalog(); var progress = new List<VbaTestResult>();
                var run = Pump(fixture.Service.RunCoverageAsync(catalog, catalog.Tests.ToArray(), progress.Add, CancellationToken.None));
                StringAssert.Contains(run.Error, "Preparation refused"); StringAssert.Contains(run.Error, "Copy close lost completion");
                Assert.AreEqual(uncertain, run.OutcomeUnknown);
                Assert.AreEqual(1, coverage.Events.Count(item => item == "Close"));
                Assert.IsNull(run.Coverage.Percent);
                Assert.IsTrue(fixture.Vbe.VBProjects.Contains(coverage.Clone));
                Assert.AreEqual(2, progress.Count);
            }
        }

        [STATestMethod]
        public void WrittenRuntimeSourceMustMatchItsReviewedPlanAfterReentrantMutation()
        {
            var project = new FakeProject();
            var code = new FakeCode();
            project.VBComponents.Add(new FakeComponent { Name = "Runtime", Type = 1, CodeModule = code });
            code.AfterAdd = () => code.Source += "\n' changed after write";
            var write = typeof(VbeTestExplorerService).GetMethod("WriteCloneModule", System.Reflection.BindingFlags.Static | System.Reflection.BindingFlags.NonPublic);
            var error = Assert.ThrowsException<System.Reflection.TargetInvocationException>(() => write.Invoke(null, new object[] { project, "Runtime", "planned source", false }));
            Assert.IsInstanceOfType(error.InnerException, typeof(InvalidOperationException));
            StringAssert.Contains(error.InnerException.Message, "did not match its plan");
            StringAssert.Contains(code.Source, "changed after write");
            Assert.AreEqual(1, code.Insertions, "An unverified mutation must never be retried.");
        }

        [STATestMethod]
        public void ActiveCoverageSessionDisposalRetainsUncertainCopyAndSettlesOnOwner()
        {
            using (var fixture = new Fixture())
            using (var coverage = new CoverageFixture(fixture))
            {
                coverage.BeforeSnapshotReturn = fixture.Service.Dispose;
                var catalog = fixture.Catalog();
                var run = Pump(fixture.Service.RunCoverageAsync(catalog, catalog.Tests.ToArray(), null, CancellationToken.None));
                Assert.IsTrue(run.OutcomeUnknown); Assert.IsFalse(run.Coverage.Available);
                Assert.IsNull(run.Coverage.Percent); Assert.IsFalse(coverage.Events.Contains("Close"));
                Assert.IsTrue(fixture.Vbe.VBProjects.Contains(coverage.Clone));
            }
        }
        [STATestMethod]
        public void CoverageHelperAcceptsAbsentProgressForSuccessfulAndRefusedCopyPreparation()
        {
            foreach (bool copyRefused in new[] { false, true })
            using (var fixture = new Fixture())
            using (var coverage = new CoverageFixture(fixture))
            {
                if (copyRefused) fixture.Service.CreateCoverageClone = (_, __, ___) => { throw new InvalidOperationException("Copy preparation refused"); };
                var catalog = fixture.Catalog();
                var helper = typeof(VbeTestExplorerService).GetMethod("ExecuteCoverageAsync", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
                var task = (System.Threading.Tasks.Task<VbaTestRun>)helper.Invoke(fixture.Service,
                    new object[] { catalog, catalog.Tests.ToArray(), null, CancellationToken.None, null });
                var run = Pump(task);
                Assert.AreEqual(2, run.Results.Count);
                Assert.AreEqual(!copyRefused, run.Coverage.Available);
                Assert.IsFalse(run.OutcomeUnknown);
                if (copyRefused) Assert.IsTrue(run.Results.All(result => result.Outcome == VbaTestOutcome.Blocked));
                else Assert.IsTrue(run.Results.All(result => result.Outcome == VbaTestOutcome.Passed));
            }
        }

        [STATestMethod]
        public void HelperReportsLateSelectionWithoutDispatchingItOutsideTheCopiedSnapshot()
        {
            using (var fixture = new Fixture())
            using (var coverage = new CoverageFixture(fixture))
            {
                var catalog = fixture.Catalog(); var tests = catalog.Tests.ToArray();
                var selection = new List<VbaTestDescriptor> { tests[0] };
                // Public BeginRun freezes its selection. This private helper additionally accepts an
                // IReadOnlyList backed by a mutable list, so its final report must account for a late item.
                coverage.AfterTest = () => selection.Add(tests[1]);
                var helper = typeof(VbeTestExplorerService).GetMethod("ExecuteCoverageAsync", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
                var task = (System.Threading.Tasks.Task<VbaTestRun>)helper.Invoke(fixture.Service,
                    new object[] { catalog, selection, null, CancellationToken.None, null });
                var run = Pump(task);
                Assert.IsNull(run.Error); Assert.AreEqual(2, run.Results.Count);
                Assert.AreEqual(VbaTestOutcome.Passed, run.Results[0].Outcome);
                Assert.AreEqual(VbaTestOutcome.Blocked, run.Results[1].Outcome);
                StringAssert.Contains(run.Results[1].Message, "stopped before this test was dispatched");
                CollectionAssert.AreEqual(new[] { "Compile", "Reset", "Test:Alpha", "Snapshot", "Close" }, coverage.Events);
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
            internal Exception CloseError;
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
                        Close = () => { Events.Add("Close"); if (CloseError != null) throw CloseError; if (CloseFails) throw new InvalidOperationException("Simulated close failure"); original.Vbe.VBProjects.Remove(copy); } };
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
