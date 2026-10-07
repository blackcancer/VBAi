using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web.Script.Serialization;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class ExcelVbeFixtureGitLayoutsPersistenceOrderTests
    {
        [DataTestMethod]
        [DataRow(false, false, "FrameMultiPage", "save")]
        [DataRow(false, true, "FrameMultiPage", "save|reopen")]
        [DataRow(true, false, "FrameMultiPage", "render|before-save|save|after-save")]
        [DataRow(true, true, "FrameMultiPage", "render|before-save|save|after-save|reopen|after-reopen")]
        [DataRow(false, false, "Image", "save")]
        [DataRow(false, true, "Image", "save|reopen")]
        [DataRow(true, false, "Image", "render|save")]
        [DataRow(true, true, "Image", "render|save|reopen")]
        public void PersistenceOrderIsExactForEveryQualifyPersistAndLayoutCombination(
            bool qualifyFonts, bool persistedBaseline, string layout, string expected)
        {
            var events = new List<string>();
            ExcelVbeFixture.PersistGitLayoutWithEvidence(qualifyFonts, persistedBaseline, layout,
                () => events.Add("render"), phase => events.Add(phase),
                () => events.Add("save"), () => events.Add("reopen"));
            CollectionAssert.AreEqual(expected.Split('|'), events.ToArray());
        }

        [TestMethod]
        public void PassiveReceiptsBracketEachOriginalPersistenceActionOnceAndKeepFailureBoundary()
        {
            var lines = new List<string>(); var actions = new List<string>();
            var failure = new IOException("Synthetic save failure");
            Action<string> visit = phase =>
            {
                var entered = new JavaScriptSerializer().Deserialize<Dictionary<string, object>>(lines.Last());
                Assert.AreEqual("Entered", entered["Boundary"], "The durable phase precedes its original action.");
                actions.Add(phase);
                if (phase == "save") throw failure;
            };
            using (NativeFixtureProgressTrace.Begin(lines.Add))
            {
                var actual = Assert.ThrowsException<IOException>(() =>
                    ExcelVbeFixture.PersistGitLayoutWithEvidence(true, true, "FrameMultiPage",
                        () => visit("render"), visit, () => visit("save"), () => visit("reopen")));
                Assert.AreSame(failure, actual);
            }
            CollectionAssert.AreEqual(new[] { "render", "before-save", "save" }, actions.ToArray());
            var boundaries = lines.Select(line => new JavaScriptSerializer().Deserialize<Dictionary<string, object>>(line))
                .Select(row => row["Phase"] + ":" + row["Boundary"]).ToArray();
            CollectionAssert.AreEqual(new[] { "Scope:Entered", "RenderDesigner:Entered", "RenderDesigner:Returned",
                "SeedBeforeSave:Entered", "SeedBeforeSave:Returned", "PersistSave:Entered", "PersistSave:Faulted", "Scope:Returned" }, boundaries);
        }

        [DataTestMethod]
        [DataRow("save")]
        [DataRow("render")]
        [DataRow("capture")]
        [DataRow("reopen")]
        public void MissingRequiredCallbackIsRejectedBeforeAnyAction(string missing)
        {
            var events = new List<string>();
            Action render = () => events.Add("render");
            Action<string> capture = phase => events.Add(phase);
            Action save = () => events.Add("save");
            Action reopen = () => events.Add("reopen");
            if (missing == "render") render = null;
            if (missing == "capture") capture = null;
            if (missing == "save") save = null;
            if (missing == "reopen") reopen = null;

            Assert.ThrowsException<ArgumentNullException>(() =>
                ExcelVbeFixture.PersistGitLayoutWithEvidence(true, true, "FrameMultiPage",
                    render, capture, save, reopen));
            Assert.AreEqual(0, events.Count, "Callback validation must finish before the first action.");
        }

        [DataTestMethod]
        [DataRow("render", "render")]
        [DataRow("before-save", "render|before-save")]
        [DataRow("save", "render|before-save|save")]
        [DataRow("after-save", "render|before-save|save|after-save")]
        [DataRow("reopen", "render|before-save|save|after-save|reopen")]
        [DataRow("after-reopen", "render|before-save|save|after-save|reopen|after-reopen")]
        public void PhaseFailurePreservesOriginalExceptionAndStopsAllLaterActions(
            string failingPhase, string expected)
        {
            var events = new List<string>();
            var failure = new IOException("Synthetic " + failingPhase + " failure");
            Action<string> visit = phase =>
            {
                events.Add(phase);
                if (phase == failingPhase) throw failure;
            };

            var actual = Assert.ThrowsException<IOException>(() =>
                ExcelVbeFixture.PersistGitLayoutWithEvidence(true, true, "FrameMultiPage",
                    () => visit("render"), visit, () => visit("save"), () => visit("reopen")));
            Assert.AreSame(failure, actual, "No wrapper or retry may replace the original failure.");
            CollectionAssert.AreEqual(expected.Split('|'), events.ToArray(),
                "A failed phase must run once and prevent every dependent phase.");
        }
    }
}
