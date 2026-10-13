using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Script.Serialization;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class VbeOptionsBatchTests
    {
        [TestMethod]
        public void CompleteBatchStagesEveryCategoryThenAcceptsOnceAndRestoresViewedCategory()
        {
            var probe = new BatchProbe(); var request = probe.Request();
            var result = Data(VbeDebugWindows.SetVbeOption(request, probe));
            Assert.AreEqual(9, probe.Writes); Assert.AreEqual(1, probe.Accepts); Assert.AreEqual(0, probe.Closes);
            Assert.AreEqual("Normal", probe.Category);
            Assert.AreEqual("Consolas", probe.Values["Font"]); Assert.AreEqual("14", probe.Values["Size"]);
            Assert.AreEqual("Off", probe.Values["Margin Indicator Bar"]);
            foreach (string category in new[] { "Normal", "Comment" })
                foreach (string property in BatchProbe.Palettes) Assert.AreEqual("Red", probe.Values[category + "." + property]);
            Assert.AreEqual(true, result["DialogClosed"]); Assert.AreEqual(false, result["PersistenceVerified"]);
        }

        [TestMethod]
        public void ImplicitPaletteQueryUsesTheOriginalCategoryAfterAnExplicitOtherCategory()
        {
            var probe = new BatchProbe(); var request = probe.Request();
            request.OptionChanges = new[] { Change("Foreground", "Red", "Comment"), Change("Foreground", "Red") };
            VbeDebugWindows.SetVbeOption(request, probe);
            Assert.AreEqual("Red", probe.Values["Normal.Foreground"]); Assert.AreEqual("Red", probe.Values["Comment.Foreground"]);
            Assert.AreEqual("Normal", probe.Category); Assert.AreEqual(1, probe.Accepts);
        }

        [TestMethod]
        public void EveryInvalidTargetIsRejectedBeforeAnyPreferenceWrite()
        {
            foreach (int scenario in Enumerable.Range(0, 12))
            {
                var probe = new BatchProbe(); var request = probe.Request();
                request.OptionChanges = new[] { Change("Font", "Consolas"), Change("Foreground", "Red", "Comment") };
                if (scenario == 0) request.ExpectedOptionsVersion = new string('0', 64);
                if (scenario == 1) request.OptionChanges[1].Pane = "Absent";
                if (scenario == 2) request.OptionChanges[1].Property = "Absent";
                if (scenario == 3) request.OptionChanges[1].Value = "Unknown";
                if (scenario == 4) request.OptionChanges[1].Query = "Absent";
                if (scenario == 5) request.OptionChanges[1] = Change("Font", "Consolas", "Comment");
                if (scenario == 6) request.OptionChanges = new[] { Change("Foreground", "Red"), Change("Foreground", "Red", "Normal") };
                if (scenario == 7) request.OptionChanges[1] = Change("Code Colors", "Comment");
                if (scenario == 8) probe.FaultProperty = "Foreground";
                if (scenario == 9) probe.DisabledProperty = "Foreground";
                if (scenario == 10) probe.AmbiguousProperty = "Foreground";
                if (scenario == 11) request.OptionChanges[1] = Change("Margin Indicator Bar", "not a boolean");
                Assert.IsNotNull(Failure(() => VbeDebugWindows.SetVbeOption(request, probe)), "scenario " + scenario);
                Assert.AreEqual(0, probe.Writes, "scenario " + scenario);
                Assert.AreEqual(0, probe.Accepts); Assert.AreEqual(1, probe.Closes);
            }
        }

        [TestMethod]
        [DataRow("absent"), DataRow("disabled"), DataRow("unreadable"), DataRow("ambiguous")]
        public void PlanningRefusalIdentifiesTargetAndCancelsWithoutPreferenceWrites(string failure)
        {
            var probe = new BatchProbe(); var request = probe.Request();
            request.OptionChanges = new[] { Change("Font", "Consolas"), Change("Foreground", "Red", "Comment") };
            int reads = 0;
            probe.OnControls = () =>
            {
                if (++reads != 4) return;
                if (failure == "absent") probe.RemovedProperty = "Foreground";
                if (failure == "disabled") probe.DisabledProperty = "Foreground";
                if (failure == "unreadable") probe.FaultProperty = "Foreground";
                if (failure == "ambiguous") probe.AmbiguousProperty = "Foreground";
            };
            var error = Failure(() => VbeDebugWindows.SetVbeOption(request, probe));
            StringAssert.Contains(error.Message, "Phase='planning'; Index=2");
            StringAssert.Contains(error.Message, "Property='" + request.OptionChanges[1].Property + "'");
            StringAssert.Contains(error.Message, "Category='Comment'");
            if (failure == "disabled") StringAssert.Contains(error.Message, "Enabled=False");
            if (failure == "unreadable") StringAssert.Contains(error.Message, "Error=unreadable");
            Assert.AreEqual(0, probe.Writes); Assert.AreEqual(0, probe.Accepts); Assert.AreEqual(1, probe.Closes);
        }

        [TestMethod]
        [DataRow("staging"), DataRow("verification")]
        public void PaletteAvailabilityChangesAfterPlanningAreDiagnosedAndNeverCommitted(string phase)
        {
            var probe = new BatchProbe(); var request = probe.Request();
            probe.OnWrite = () =>
            {
                if (phase == "staging" && probe.Writes == 2) probe.DisabledProperty = "Foreground";
                if (phase == "verification" && probe.Writes == 9) probe.DisabledProperty = "Foreground";
            };
            var error = Failure(() => VbeDebugWindows.SetVbeOption(request, probe));
            StringAssert.Contains(error.Message, "Phase='" + phase + "'");
            StringAssert.Contains(error.Message, "Property='Foreground'");
            StringAssert.Contains(error.Message, "Enabled=False");
            Assert.AreEqual(phase == "staging" ? 2 : 9, probe.Writes);
            Assert.AreEqual(0, probe.Accepts); Assert.AreEqual(1, probe.Closes);
        }

        [TestMethod]
        public void MalformedBatchStructureDoesNotEvenOpenOrInspectNativeOptions()
        {
            foreach (int scenario in Enumerable.Range(0, 9))
            {
                var probe = new BatchProbe(); var request = probe.Request(); probe.Reads = 0;
                if (scenario == 0) request = null;
                if (scenario == 1) request.OptionChanges = null;
                if (scenario == 2) request.OptionChanges = new VbeOptionChange[0];
                if (scenario == 3) request.OptionChanges = Enumerable.Repeat(Change("Font", "Consolas"), 65).ToArray();
                if (scenario == 4) request.OptionChanges[0] = null;
                if (scenario == 5) request.OptionChanges[0].Pane = " ";
                if (scenario == 6) request.Pane = "Editor Format";
                if (scenario == 7) request.Value = false;
                if (scenario == 8) request.ExpectedOptionsVersion = " ";
                Assert.ThrowsException<ArgumentException>(() => VbeDebugWindows.SetVbeOptionBatch(request, probe));
                Assert.AreEqual(0, probe.Reads); Assert.AreEqual(0, probe.Writes); Assert.AreEqual(0, probe.Closes);
            }
        }

        [TestMethod]
        public void WriteAndCommitExceptionsNeverTriggerCancellationOrMutationReplay()
        {
            foreach (bool atCommit in new[] { false, true })
            {
                var probe = new BatchProbe(); var request = probe.Request();
                if (atCommit) probe.OnAccept = () => throw new InvalidOperationException("uncertain OK");
                else probe.OnWrite = () => { if (probe.Writes == 2) throw new InvalidOperationException("uncertain write"); };
                Assert.IsNotNull(Failure(() => VbeDebugWindows.SetVbeOption(request, probe)));
                Assert.AreEqual(atCommit ? 9 : 2, probe.Writes);
                Assert.AreEqual(atCommit ? 1 : 0, probe.Accepts); Assert.AreEqual(0, probe.Closes);
            }
        }

        [TestMethod]
        public void FinalVerificationDetectsLaterChangesToEarlierValuesBeforeOneCancellation()
        {
            var probe = new BatchProbe(); var request = probe.Request();
            probe.OnWrite = () => { if (probe.Writes == 9) probe.Values["Font"] = "Courier New"; };
            Assert.IsNotNull(Failure(() => VbeDebugWindows.SetVbeOption(request, probe)));
            Assert.AreEqual(9, probe.Writes); Assert.AreEqual(0, probe.Accepts); Assert.AreEqual(1, probe.Closes);
        }

        [TestMethod]
        public void OpenDialogAfterOkIsReportedAsUnverifiedClosureWithoutRetry()
        {
            var probe = new BatchProbe { KeepOpen = true }; var request = probe.Request();
            var result = Data(VbeDebugWindows.SetVbeOption(request, probe));
            Assert.AreEqual(false, result["DialogClosed"]); Assert.AreEqual(1, probe.Accepts); Assert.AreEqual(0, probe.Closes);
        }

        [TestMethod]
        public void UnretainedOriginalCategoryPreventsCommitAndRequestsCancellationOnce()
        {
            var probe = new BatchProbe(); var request = probe.Request(); int originalSelections = 0;
            probe.RefuseSelection = category => probe.Writes == 9 && category == "Normal" && ++originalSelections == 4;
            Assert.IsNotNull(Failure(() => VbeDebugWindows.SetVbeOption(request, probe)));
            Assert.AreEqual(9, probe.Writes); Assert.AreEqual(0, probe.Accepts); Assert.AreEqual(1, probe.Closes);
        }

        [TestMethod]
        public void RequestTargetsAndRevisionAreFrozenBeforeTheFirstNativeWait()
        {
            var probe = new BatchProbe(); var request = probe.Request();
            probe.OnPause = () => { request.OptionChanges[0].Value = "Unknown"; request.ExpectedOptionsVersion = "changed"; };
            VbeDebugWindows.SetVbeOption(request, probe);
            Assert.AreEqual("Consolas", probe.Values["Font"]); Assert.AreEqual(1, probe.Accepts);
        }

        private static VbeOptionChange Change(string property, object value, string category = null) =>
            new VbeOptionChange { Pane = "Editor Format", Property = property, Value = value, Query = category };
        private static IDictionary<string, object> Data(object value)
        { var json = new JavaScriptSerializer(); return json.Deserialize<Dictionary<string, object>>(json.Serialize(value)); }
        private static Exception Failure(Action action)
        { try { action(); } catch (Exception error) { return error; } Assert.Fail("Failure expected."); return null; }

        /// <summary>Models independent native category values; production validation, planning and commit orchestration remain real.</summary>
        private sealed class BatchProbe : VbeDebugWindows.IWritableOptionsProbe, VbeDebugWindows.IFormatCategoriesOptionsProbe
        {
            internal static readonly string[] Palettes = { "Foreground", "Background", "Indicator" };
            internal readonly Dictionary<string, object> Values = new Dictionary<string, object>
                { ["Font"] = "Courier New", ["Size"] = "12", ["Margin Indicator Bar"] = "On" };
            internal string Category = "Normal", FaultProperty, DisabledProperty, AmbiguousProperty, RemovedProperty;
            internal int Reads, Writes, Accepts, Closes;
            internal bool Open = true, KeepOpen;
            internal Action OnWrite, OnAccept, OnPause, OnControls;
            internal Func<string, bool> RefuseSelection;
            internal BatchProbe()
            { foreach (string category in new[] { "Normal", "Comment" }) foreach (string name in Palettes) Values[category + "." + name] = "Black"; }
            public IntPtr Dialog() => Open ? new IntPtr(1) : IntPtr.Zero;
            public IList<string> Tabs(IntPtr dialog) { Reads++; return new[] { "Editor Format" }; }
            private VbeDebugWindows.OptionsControl Control(string name, string type, object value, params string[] choices) =>
                new VbeDebugWindows.OptionsControl { Name = name, Type = type, Value = value, Choices = choices,
                    Error = name == FaultProperty ? "unreadable" : null, Enabled = name != DisabledProperty };
            private IList<VbeDebugWindows.OptionsControl> PaletteControls(string category) => Palettes.Select(name =>
                Control(name, "ControlType.ComboBox", Values[category + "." + name], "Black", "Red")).ToArray();
            public IList<VbeDebugWindows.OptionsControl> Controls(IntPtr dialog, int tabIndex)
            {
                OnControls?.Invoke();
                Reads++;
                var controls = new List<VbeDebugWindows.OptionsControl> {
                    Control("Font", "ControlType.ComboBox", Values["Font"], "Courier New", "Consolas"),
                    Control("Size", "ControlType.ComboBox", Values["Size"], "12", "14"),
                    Control("Margin Indicator Bar", "ControlType.CheckBox", Values["Margin Indicator Bar"]),
                    Control("Code Colors", "ControlType.List", Category, "Normal", "Comment") };
                controls.AddRange(PaletteControls(Category));
                if (RemovedProperty != null) controls.RemoveAll(control => control.Name == RemovedProperty);
                if (AmbiguousProperty != null) controls.Add(controls.Single(control => control.Name == AmbiguousProperty));
                return controls;
            }
            public IList<VbeDebugWindows.OptionsFormatCategory> FormatCategories(IntPtr dialog, int tabIndex) =>
                new[] { "Normal", "Comment" }.Select(category => new VbeDebugWindows.OptionsFormatCategory { Category = category, Palettes = PaletteControls(category) }).ToArray();
            public void SelectFormatCategory(IntPtr dialog, int tabIndex, string category)
            { if (category != "Normal" && category != "Comment") throw new InvalidOperationException("Unknown category."); if (RefuseSelection?.Invoke(category) != true) Category = category; }
            public IList<VbeDebugWindows.OptionsChoice> ErrorChoices(IntPtr dialog) => new VbeDebugWindows.OptionsChoice[0];
            public void Pause(int milliseconds) => OnPause?.Invoke();
            public void Close(IntPtr dialog) { Closes++; Open = false; }
            public void Accept(IntPtr dialog) { Accepts++; OnAccept?.Invoke(); if (!KeepOpen) Open = false; }
            public void Write(IntPtr dialog, int tabIndex, string name, string type, object value)
            {
                Writes++;
                string key = Palettes.Contains(name) ? Category + "." + name : name;
                Values[key] = value is bool check ? (object)(check ? "On" : "Off") : value;
                OnWrite?.Invoke();
            }
            internal Request Request()
            {
                dynamic read = VbeDebugWindows.ReadVbeOptions(this); Open = true; Closes = 0;
                var changes = new List<VbeOptionChange> { Change("Font", "Consolas"), Change("Size", "14") };
                foreach (string category in new[] { "Normal", "Comment" }) foreach (string name in Palettes) changes.Add(Change(name, "Red", category));
                changes.Add(Change("Margin Indicator Bar", false));
                return new Request { ExpectedOptionsVersion = read.OptionsVersion, OptionChanges = changes.ToArray() };
            }
        }
    }
}
