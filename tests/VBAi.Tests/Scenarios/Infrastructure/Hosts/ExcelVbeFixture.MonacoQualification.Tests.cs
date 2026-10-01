using System;
using System.Collections.Generic;
using System.Globalization;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    /// <summary>Validates the actual-host observation gates with scalar data only; these tests are never native acceptance.</summary>
    [TestClass, TestCategory("Unit")]
    public sealed class ExcelNativeMonacoObservationTests
    {
        private const string Closed = "VBAProject · ModuleClosedScope", Live = "VBAProject · ModuleLiveScope";

        [TestMethod]
        public void DecoratedInstalledStatusDoesNotMislabelAccessibilityChromeAsRenderedSource()
        {
            var actualShape = Valid();
            actualShape.RenderedText = "Éditeur VBAi\nVBAProject · ModuleLiveScope\nVBAProject · ModuleClosedScope\nContenu de l'éditeur\n￼\nVBA editor";
            actualShape.TextProviderProcessIds = new[] { 123, 456, 789 };
            ExcelVbeFixture.RequireInstalledMonacoObservation(actualShape, 123, Closed, Live, ExcelVbeFixture.MonacoSynchronizedStatus);
            Assert.AreEqual("LIVE_SOURCE_NOT_EXPOSED_BY_ACCESSIBILITY", ExcelVbeFixture.MonacoAccessibleSourceStatus(actualShape));
        }

        [TestMethod]
        public void ActualEmbeddedOwnedLiveObservationAcceptsRetainedClosedTabAndBrowserProviderPid()
        {
            var observation = Valid();
            observation.TextProviderProcessIds = new[] { 123, 456 };
            ExcelVbeFixture.RequireInstalledMonacoObservation(observation, 123, Closed, Live, ExcelVbeFixture.MonacoSynchronizedStatus);
        }

        [TestMethod]
        public void NativeWinFormsLabelCanBindWhenUiaProxyDoesNotExposeItsDesignerName()
        {
            var observation = Valid();
            observation.StatusAutomationId = "NativeProxyId";
            observation.StatusClassName = "WindowsForms10.STATIC.app.fake";
            ExcelVbeFixture.RequireInstalledMonacoObservation(observation, 123, Closed, Live, ExcelVbeFixture.MonacoSynchronizedStatus);
            observation.StatusClassName = "BrowserText";
            Assert.ThrowsException<AssertFailedException>(() => ExcelVbeFixture.RequireInstalledMonacoObservation(observation, 123, Closed, Live, ExcelVbeFixture.MonacoSynchronizedStatus));
        }

        [DataTestMethod]
        [DataRow("no-editor"), DataRow("detached"), DataRow("vbe-owner"), DataRow("editor-owner"), DataRow("same-handle")]
        [DataRow("hidden"), DataRow("class"), DataRow("status-owner"), DataRow("status-hidden"), DataRow("status-id")]
        [DataRow("duplicate-status"), DataRow("warning"), DataRow("pending"), DataRow("missing-closed-tab"), DataRow("missing-live-tab")]
        [DataRow("wrong-tab-owner"), DataRow("old-selected"), DataRow("two-selected"), DataRow("foreign-caption"), DataRow("duplicate-caption")]
        [DataRow("truncated"), DataRow("provider-error")]
        public void ForeignDetachedStaleOrUnrenderedObservationCannotQualifyNativeMonaco(string fault)
        {
            var value = Valid();
            switch (fault)
            {
                case "no-editor": value.EditorHandle = 0; break;
                case "detached": value.Embedded = false; break;
                case "vbe-owner": value.VbeProcessId = 456; break;
                case "editor-owner": value.EditorProcessId = 456; break;
                case "same-handle": value.EditorHandle = value.VbeHandle; break;
                case "hidden": value.Visible = false; break;
                case "class": value.EditorClass = "NativeCodePane"; break;
                case "status-owner": value.StatusProcessId = 456; break;
                case "status-hidden": value.StatusVisible = false; break;
                case "status-id": value.StatusAutomationId = "OtherStatus"; break;
                case "duplicate-status": value.StatusCount = 2; break;
                case "warning": value.StatusText = ExcelVbeFixture.MonacoClosedStatus; break;
                case "pending": value.StatusText = "Changes pending synchronization with VBA."; break;
                case "missing-closed-tab": value.Tabs = new[] { Live }; value.TabProcessIds = new[] { 123 }; break;
                case "missing-live-tab": value.Tabs = new[] { Closed }; value.TabProcessIds = new[] { 123 }; break;
                case "wrong-tab-owner": value.TabProcessIds = new[] { 123, 456 }; break;
                case "old-selected": value.SelectedTabs = new[] { Closed }; break;
                case "two-selected": value.SelectedTabs = new[] { Closed, Live }; break;
                case "foreign-caption": value.Tabs = new[] { Closed, "OtherProject · ModuleLiveScope" }; break;
                case "duplicate-caption": value.Tabs = new[] { Closed, Live, Live }; value.TabProcessIds = new[] { 123, 123, 123 }; break;
                case "truncated": value.TreeTruncated = true; break;
                case "provider-error": value.ObservationError = "The provider became unavailable."; break;
            }
            Assert.ThrowsException<AssertFailedException>(() => ExcelVbeFixture.RequireInstalledMonacoObservation(value, 123, Closed, Live, ExcelVbeFixture.MonacoSynchronizedStatus));
        }

        [DataTestMethod, DataRow("live"), DataRow("empty"), DataRow("old"), DataRow("chrome")]
        public void AccessibilityAssessmentDoesNotPromoteEmptyOldOrChromeTextToLiveSourceProof(string shape)
        {
            var value = Valid();
            if (shape == "empty") value.RenderedText = null;
            if (shape == "old") value.RenderedText = "Option Explicit\nClosedValue = 17\nVBAi owned closed scope probe";
            if (shape == "chrome") value.RenderedText = "VBAProject · ModuleLiveScope\nContenu de l'éditeur\nVBA editor";
            Assert.AreEqual(shape == "live" ? "LIVE_SOURCE_MARKERS_OBSERVED" : "LIVE_SOURCE_NOT_EXPOSED_BY_ACCESSIBILITY",
                ExcelVbeFixture.MonacoAccessibleSourceStatus(value));
            Assert.AreEqual(Live, ExcelVbeFixture.MonacoTabCaption("VBAProject", "ModuleLiveScope"));
        }

        [TestMethod]
        public void FrenchNativeStatusRequiresFrenchCatalogueTextAndStillRejectsClosedWarning()
        {
            var culture = CultureInfo.GetCultureInfo("fr-FR");
            string expected = ExcelVbeFixture.MonacoLocalized(ExcelVbeFixture.MonacoSynchronizedStatus, culture);
            Assert.AreNotEqual(ExcelVbeFixture.MonacoSynchronizedStatus, expected);
            var value = Valid(); value.StatusText = expected;
            ExcelVbeFixture.RequireInstalledMonacoObservation(value, 123, Closed, Live, expected);
            value.StatusText = ExcelVbeFixture.MonacoLocalized(ExcelVbeFixture.MonacoClosedStatus, culture);
            Assert.ThrowsException<AssertFailedException>(() => ExcelVbeFixture.RequireInstalledMonacoObservation(value, 123, Closed, Live, expected));
        }

        [DataTestMethod, DataRow("valid"), DataRow("empty"), DataRow("test-payload"), DataRow("host-payload"), DataRow("owner")]
        public void NativeCandidateRequiresExactFrozenPayloadInTestReferenceAndOwningBridge(string scenario)
        {
            var expected = Guid.Parse("f99a5c32-17e4-4a92-9d4c-22a973e515b7");
            var reference = expected;
            var status = new Dictionary<string, object> { ["HostProcessId"] = 123, ["AssemblyModuleVersionId"] = expected.ToString("D") };
            if (scenario == "empty") expected = Guid.Empty;
            if (scenario == "test-payload") reference = Guid.NewGuid();
            if (scenario == "host-payload") status["AssemblyModuleVersionId"] = Guid.NewGuid().ToString("D");
            if (scenario == "owner") status["HostProcessId"] = 456;
            if (scenario == "valid") ExcelVbeFixture.RequireMonacoCandidate(expected, reference, 123, status);
            else Assert.ThrowsException<AssertFailedException>(() => ExcelVbeFixture.RequireMonacoCandidate(expected, reference, 123, status));
        }

        private static ExcelVbeFixture.MonacoNativeObservation Valid() => new ExcelVbeFixture.MonacoNativeObservation {
            VbeHandle = 100, VbeProcessId = 123, EditorHandle = 200, EditorProcessId = 123, EditorClass = "WindowsForms10.Window.fake",
            Embedded = true, Visible = true, StatusProcessId = 123, StatusVisible = true, StatusAutomationId = "status", StatusCount = 1,
            StatusText = ExcelVbeFixture.MonacoSynchronizedStatus, Tabs = new[] { Closed, Live }, TabProcessIds = new[] { 123, 123 },
            SelectedTabs = new[] { Live }, RenderedText = "Option Explicit\n' VBAi owned live scope probe\nPublic Function LiveValue() As Long\nLiveValue = 42\nEnd Function"
        };
    }
}
