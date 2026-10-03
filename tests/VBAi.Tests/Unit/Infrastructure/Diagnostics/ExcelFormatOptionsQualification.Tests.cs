using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Web.Script.Serialization;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class ExcelFormatOptionsQualificationTests
    {
        [TestMethod]
        public void HistoricalPrefixKeepsNullQueryAndRestoresOnlyConfirmedFontAndPaletteWrites()
        {
            var probe = new Probe { EmptySizes = true };
            probe.Create(historicalPalettePrefix: true).Run();
            Assert.AreEqual(9, probe.Writes, "Four real commits, four compensations and one expected size refusal are required.");
            Assert.AreEqual(1, probe.Cleanup);
            Assert.IsFalse(probe.Phases.Contains("MarginIntent"));
            Assert.IsFalse(probe.Phases.Contains("OtherCategoryIntent"));
            Assert.IsFalse(probe.Phases.Contains("StaleVersionIntent"));
            foreach (var item in probe.Requests.Where(request => Equals(request["Command"], "set_vbe_option")))
                if (new[] { "Foreground", "Background", "Indicator" }.Contains(Convert.ToString(item["Property"])))
                    Assert.IsNull(item["Query"], "The historical request omitted Query, including restoration.");
            Assert.IsTrue(probe.Phases.Contains("BaselineRestored"));
        }

        [TestMethod]
        public void FocusedMarginQualificationChangesAndRestoresOnlyTheRealCheckbox()
        {
            var probe = new Probe();
            probe.Create(verifyReadStability: true, marginOnly: true).Run();
            Assert.AreEqual(2, probe.Writes, "One transition and its single restoration are required.");
            Assert.AreEqual("On", probe.Values["Margin Indicator Bar"]);
            Assert.AreEqual(1, probe.Cleanup);
            Assert.IsFalse(probe.Phases.Contains("FontIntent"));
            Assert.IsFalse(probe.Phases.Contains("StaleVersionIntent"));
            Assert.IsTrue(probe.Phases.Contains("BaselineRestored"));
        }

        [TestMethod]
        public void ChangedCompleteTabsCannotPassRestorationWithAnUnchangedRevision()
        {
            var probe = new Probe { SemanticFault = "CompleteRestorationReadbackIntent" };
            var runner = probe.Create();
            Assert.IsNotNull(Failure(runner.Run));
            Assert.IsTrue(runner.HostRetained);
            Assert.AreEqual(0, probe.Cleanup);
            Assert.IsFalse(probe.Phases.Contains("BaselineRestored"));
        }

        private sealed class Probe
        {
            internal bool EmptySizes, CleanupFault, PreservationFault;
            internal int FaultAt, Reads, Writes, Cleanup, Preserved, ClosureCalls;
            internal string Fault, EvidenceFault, SemanticFault, ClosureFault;
            internal readonly List<string> Phases = new List<string>();
            internal readonly List<IDictionary<string, object>> Requests = new List<IDictionary<string, object>>();
            internal readonly List<Tuple<string, object>> Evidence = new List<Tuple<string, object>>();
            internal readonly Dictionary<string, object> Values = new Dictionary<string, object> {
                ["Font"] = "Courier New", ["Size"] = "12", ["Margin Indicator Bar"] = "On",
                ["Normal Text.Foreground"] = "Black", ["Normal Text.Background"] = "White", ["Normal Text.Indicator"] = "Blue",
                ["Keyword Text.Foreground"] = "Black", ["Keyword Text.Background"] = "White", ["Keyword Text.Indicator"] = "Blue" };
            internal string Category = "Normal Text", BaselineVersion;
            internal ExcelFormatOptionsQualification Create(bool verifyReadStability = false, bool marginOnly = false, bool historicalPalettePrefix = false)
            {
                BaselineVersion = Version();
                return new ExcelFormatOptionsQualification(42, Dispatch, ObserveClosure,
                    () => { Preserved++; if (PreservationFault) throw new IOException("retention failed"); },
                    () => { Cleanup++; if (CleanupFault) throw new IOException("cleanup failed"); },
                    (phase, data) => {
                        Phases.Add(phase); Evidence.Add(Tuple.Create(phase, data));
                        if (EvidenceFault == phase) { EvidenceFault = null; throw new IOException("evidence failed at " + phase); }
                    }, verifyReadStability, marginOnly, historicalPalettePrefix);
            }
            private IDictionary<string, object> Dispatch(object raw)
            {
                string phase = Phases.Last();
                StringAssert.EndsWith(phase, "Intent", "Durable intent must precede every dispatch.");
                var request = Copy(raw); Requests.Add(request);
                IDictionary<string, object> response;
                if (Equals(request["Command"], "read_vbe_options"))
                {
                    Reads++;
                    if (Fault == "AlreadyBaseline" && phase == "BeforeRestorationIntent") Values["Font"] = "Courier New";
                    var state = State();
                    if (SemanticFault == phase || (Fault == "RestorationMismatch" && phase == "RestorationReadbackIntent")) { SemanticFault = null;
                        var format = (IDictionary<string, object>)((object[])state["Tabs"])[0];
                        ((IDictionary<string, object>)((object[])format["Controls"])[0])["Value"] = "unrequested font";
                    }
                    if (Fault == "BaselineMismatch" && phase == "CompleteRestorationReadbackIntent") state["OptionsVersion"] = new string('e', 64);
                    if (Fault == "RefusalChangedRevision" && phase.EndsWith("ClosedReadbackIntent")) state["OptionsVersion"] = new string('e', 64);
                    response = Reply(state);
                }
                else
                {
                    Writes++;
                    if (Equals(request["ExpectedOptionsVersion"], BaselineVersion) && phase == "StaleVersionIntent")
                        response = Refused("VBE options changed since inspection; read them again.");
                    else
                    {
                        Assert.AreEqual(Version(), request["ExpectedOptionsVersion"], "Each real write must guard the exact fresh complete revision.");
                        if (EmptySizes && Equals(request["Property"], "Size"))
                            response = Refused("The exact native choice is absent, ambiguous or unreadable.");
                        else
                        {
                            string property = (string)request["Property"];
                            string query = request.TryGetValue("Query", out object q) ? q as string : null;
                            if (!string.IsNullOrEmpty(query)) Category = query;
                            string key = new[] { "Foreground", "Background", "Indicator" }.Contains(property) ? Category + "." + property : property;
                            Values[key] = request["Value"] is bool check ? (object)(check ? "On" : "Off") : request["Value"];
                            response = Reply(new Dictionary<string, object> { ["CommitRequested"] = true, ["ControlValueVerified"] = true, ["DialogClosed"] = true });
                        }
                    }
                }
                if (Requests.Count != FaultAt) return response;
                if (Fault == "Lost") throw new IOException("native response lost after emission");
                if (Fault == "Timeout") throw new TimeoutException("native reply timed out after emission");
                if (Fault == "Null") return null;
                if (Fault == "Refused") return Refused("unexpected native refusal");
                if (Fault == "MissingOk") { response.Remove("Ok"); return response; }
                if (Fault == "NoData") { response["Ok"] = true; response["Data"] = null; return response; }
                if (Fault == "TopPending") { response["Pending"] = true; return response; }
                if (Fault == "WrongExpectedError") { response["Error"] = "different validation refusal"; return response; }
                if (Fault == "RefusalData") { response["Data"] = new Dictionary<string, object> { ["DialogClosed"] = true }; return response; }
                if (Fault == "UnexpectedSuccess") return Reply(new Dictionary<string, object> { ["DialogClosed"] = true, ["CommitRequested"] = true, ["ControlValueVerified"] = true });
                var data = response["Data"] as IDictionary<string, object>;
                if (data == null) { response["Pending"] = true; return response; }
                if (Fault == "OpenDialog") data["DialogClosed"] = false;
                else if (Fault == "NoCommit") data["CommitRequested"] = false;
                else if (Fault == "NoVerification") data["ControlValueVerified"] = false;
                else if (Fault == "BadRevision") data["OptionsVersion"] = "not a sha";
                else data[Fault] = true;
                return response;
            }
            private IDictionary<string, object> ObserveClosure()
            {
                ClosureCalls++;
                if (ClosureFault == "Throw") throw new IOException("native enumeration failed");
                if (ClosureFault == "Null") return null;
                var observation = new Dictionary<string, object> { ["ProcessId"] = 42, ["ProcessIdentityVerified"] = true,
                    ["EnumerationSucceeded"] = true, ["OptionsDialogAbsent"] = true, ["ObservationOnly"] = true };
                if (ClosureFault == "ForeignPid") observation["ProcessId"] = 43;
                else if (ClosureFault != null) observation[ClosureFault] = false;
                return observation;
            }
            internal string Version()
            {
                using (var hash = SHA256.Create()) return BitConverter.ToString(hash.ComputeHash(Encoding.UTF8.GetBytes(
                    new JavaScriptSerializer().Serialize(new { Values, Category })))).Replace("-", "").ToLowerInvariant();
            }
            private IDictionary<string, object> State() => new Dictionary<string, object> {
                ["OptionsVersion"] = Version(), ["DialogClosed"] = true,
                ["Tabs"] = new object[] { new Dictionary<string, object> { ["Tab"] = "Editor Format",
                    ["Controls"] = new object[] {
                        Control("Font", "ControlType.ComboBox", Values["Font"], "Consolas", "Courier New"),
                        Control("Size", "ControlType.ComboBox", Values["Size"], EmptySizes ? new string[0] : new[] { "12", "14" }),
                        Control("Margin Indicator Bar", "ControlType.CheckBox", Values["Margin Indicator Bar"]),
                        Control("Code Colors", "ControlType.List", Category, "Normal Text", "Keyword Text"),
                        Control("Foreground", "ControlType.ComboBox", Values[Category + ".Foreground"], "Black", "Red"),
                        Control("Background", "ControlType.ComboBox", Values[Category + ".Background"], "White", "Yellow"),
                        Control("Indicator", "ControlType.ComboBox", Values[Category + ".Indicator"], "Blue", "Green") },
                    ["FormatCategories"] = new object[] { CategoryState("Normal Text"), CategoryState("Keyword Text") } } } };
            private IDictionary<string, object> CategoryState(string category) => new Dictionary<string, object> { ["Category"] = category,
                ["Palettes"] = new object[] { Control("Foreground", "ControlType.ComboBox", Values[category + ".Foreground"], "Black", "Red"),
                    Control("Background", "ControlType.ComboBox", Values[category + ".Background"], "White", "Yellow"),
                    Control("Indicator", "ControlType.ComboBox", Values[category + ".Indicator"], "Blue", "Green") } };
        }
        private static IDictionary<string, object> Control(string name, string type, object value, params string[] choices) =>
            new Dictionary<string, object> { ["Name"] = name, ["Type"] = type, ["Value"] = value, ["Choices"] = choices.Cast<object>().ToArray(), ["Error"] = null };
        private static IDictionary<string, object> Reply(object data) => new Dictionary<string, object> { ["Ok"] = true, ["Data"] = data };
        private static IDictionary<string, object> Refused(string error) => new Dictionary<string, object> { ["Ok"] = false, ["Error"] = error, ["Data"] = null };
        private static Dictionary<string, object> Copy(object value) { var json = new JavaScriptSerializer(); return json.Deserialize<Dictionary<string, object>>(json.Serialize(value)); }
        private static Exception Failure(Action action) { try { action(); Assert.Fail("The original failure must not become acceptance."); }
            catch (AssertFailedException error) when (error.Message == "The original failure must not become acceptance.") { throw; }
            catch (Exception error) { return error; } return null; }

        [TestMethod]
        public void DisabledOwnedFormatBootstrapSkipsBeforeValidatingPathsPreparingFilesOrDispatching()
        {
            int prepared = 0, started = 0, qualified = 0;
            Assert.ThrowsException<AssertInconclusiveException>(() => ExcelFormatOptionsQualification.RunOwned<object>(false,
                null, null, "inherited-manifest", () => prepared++, trace => { started++; return new object(); }, host => qualified++));
            Assert.AreEqual(0, prepared); Assert.AreEqual(0, started); Assert.AreEqual(0, qualified);
        }

        [DataTestMethod, DataRow(null), DataRow("relative"), DataRow("C:relative"), DataRow("\\\\server\\share\\results"), DataRow("C:\\results:stream")]
        public void OwnedResultsMustBeAbsoluteAndLocalBeforeAnyEvidencePreparationOrStartup(string results)
        {
            int prepared = 0, started = 0, qualified = 0;
            Assert.ThrowsException<ArgumentException>(() => ExcelFormatOptionsQualification.RunOwned<object>(true,
                results, Path.GetTempPath(), null, () => prepared++, trace => { started++; return new object(); }, host => qualified++));
            Assert.AreEqual(0, prepared); Assert.AreEqual(0, started); Assert.AreEqual(0, qualified);
        }

        [DataTestMethod, DataRow(null), DataRow("relative"), DataRow("C:relative"), DataRow("\\\\server\\share\\evidence"), DataRow("C:\\evidence:stream")]
        public void EvidenceRootMustBeAbsoluteAndLocalBeforeAnyEvidencePreparationOrStartup(string evidence)
        {
            int prepared = 0, started = 0, qualified = 0;
            Assert.ThrowsException<ArgumentException>(() => ExcelFormatOptionsQualification.RunOwned<object>(true,
                Path.GetTempPath(), evidence, null, () => prepared++, trace => { started++; return new object(); }, host => qualified++));
            Assert.AreEqual(0, prepared); Assert.AreEqual(0, started); Assert.AreEqual(0, qualified);
        }

        [DataTestMethod, DataRow("manifest.json"), DataRow(" ")]
        public void InheritedPathVisibilityManifestIsRejectedWithoutReadingItOrLaunching(string manifest)
        {
            int prepared = 0, started = 0, qualified = 0;
            StringAssert.Contains(Failure(() => ExcelFormatOptionsQualification.RunOwned<object>(true, Path.GetTempPath(),
                Path.GetTempPath(), manifest, () => prepared++, trace => { started++; return new object(); }, host => qualified++)).Message, "must not inherit");
            Assert.AreEqual(0, prepared); Assert.AreEqual(0, started); Assert.AreEqual(0, qualified);
        }

        [TestMethod]
        public void EvidencePreparationFailureNeverEntersOwnedStartupOrOptionsLifecycle()
        {
            int started = 0, qualified = 0; var original = new IOException("evidence root unavailable");
            var failure = Failure(() => ExcelFormatOptionsQualification.RunOwned<object>(true, Path.GetTempPath(), Path.GetTempPath(), null,
                () => { throw original; }, trace => { started++; return new object(); }, host => qualified++));
            Assert.AreSame(original, failure); Assert.AreEqual(0, started); Assert.AreEqual(0, qualified);
        }

        [DataTestMethod, DataRow("Timeout"), DataRow("IO"), DataRow("AssemblyIdentity")]
        public void StartupStatusFailureNeverDispatchesFormatOptionsOrRunsCleanupAndKeepsOriginalFailure(string kind)
        {
            var probe = new Probe(); int prepared = 0, starts = 0, statusCalls = 0, qualified = 0;
            Exception original = kind == "Timeout" ? (Exception)new TimeoutException("startup status pending") : kind == "IO"
                ? (Exception)new IOException("startup status response lost") : new InvalidOperationException("loaded assembly identity differs");
            var failure = Failure(() => ExcelFormatOptionsQualification.RunOwned<Probe>(true, Path.GetTempPath(), Path.GetTempPath(), null,
                () => prepared++, trace => { starts++; statusCalls++; throw original; }, host => { qualified++; host.Create().Run(); }));
            Assert.AreSame(original, failure); Assert.AreEqual(1, prepared); Assert.AreEqual(1, starts); Assert.AreEqual(1, statusCalls);
            Assert.AreEqual(0, qualified); Assert.AreEqual(0, probe.Requests.Count); Assert.AreEqual(0, probe.Cleanup);
        }

        [TestMethod]
        public void MissingReadyFixtureNeverConstructsTheOptionsLifecycleOrCleanup()
        {
            var probe = new Probe(); int starts = 0, qualified = 0;
            StringAssert.Contains(Failure(() => ExcelFormatOptionsQualification.RunOwned<Probe>(true, Path.GetTempPath(), Path.GetTempPath(), null,
                () => { }, trace => { starts++; return null; }, host => { qualified++; host.Create().Run(); })).Message, "no ready fixture");
            Assert.AreEqual(1, starts); Assert.AreEqual(0, qualified); Assert.AreEqual(0, probe.Requests.Count); Assert.AreEqual(0, probe.Cleanup);
        }

        [DataTestMethod, DataRow(null), DataRow("")]
        public void OneOwnedReadyBootstrapDelegatesOnceToFullMatrixWithFreshLocalTraceAndNoManifest(string manifest)
        {
            var probe = new Probe(); int prepared = 0, starts = 0, qualified = 0;
            string root = Path.Combine(Path.GetTempPath(), "VBAi-format-contract-" + Guid.NewGuid().ToString("N"));
            ExcelFormatOptionsQualification.RunOwned(true, Path.GetTempPath(), root, manifest, () => prepared++, trace => {
                starts++; Assert.AreEqual(1, prepared); Assert.AreEqual(Path.Combine(Path.GetFullPath(root), "owned-bootstrap-phases.jsonl"), trace);
                return probe;
            }, host => { qualified++; Assert.AreSame(probe, host); host.Create().Run(); });
            Assert.AreEqual(1, starts); Assert.AreEqual(1, qualified); Assert.AreEqual(1, probe.Cleanup);
            Assert.AreEqual(probe.BaselineVersion, probe.Version());
        }

        [DataTestMethod, DataRow(false), DataRow(true)]
        public void CompleteNineScenarioMatrixPreservesAllCategoriesAndRestoresBaseline(bool emptySizes)
        {
            var probe = new Probe { EmptySizes = emptySizes }; var runner = probe.Create(); runner.Run();
            Assert.AreEqual(9, ExcelFormatOptionsQualification.Scenarios.Length);
            Assert.AreEqual(probe.BaselineVersion, probe.Version()); Assert.AreEqual("Normal Text", probe.Category);
            Assert.AreEqual(1, probe.Cleanup); Assert.AreEqual(0, probe.Preserved); Assert.IsFalse(runner.HostRetained);
            Assert.AreEqual(emptySizes ? 2 : 1, probe.ClosureCalls);
            Assert.AreEqual(1, probe.Phases.Count(phase => phase == "BaselineRestored"));
            foreach (string scenario in new[] { "FontVerified", "ForegroundVerified", "BackgroundVerified", "IndicatorVerified", "OtherCategoryVerified", "MarginVerified", "StaleVersionVerifiedRefusal" })
                Assert.IsTrue(probe.Phases.Contains(scenario), scenario);
            Assert.IsTrue(probe.Phases.Contains(emptySizes ? "EmptySizeVerifiedRefusal" : "SizeVerified"));
            var actualRestores = probe.Requests.Where(request => request.ContainsKey("Query") && Equals(request["Command"], "set_vbe_option")).ToArray();
            Assert.IsTrue(actualRestores.Any(item => Equals(item["Query"], "Normal Text")));
            Assert.IsTrue(actualRestores.Any(item => Equals(item["Query"], "Keyword Text")));
            Assert.IsTrue(probe.Evidence.Any(item => item.Item1 == "BaselineComplete"));
            var terminal = Copy(probe.Evidence.Single(item => item.Item1 == "QualificationTerminal").Item2);
            Assert.AreEqual(true, terminal["Verified"]); Assert.AreEqual(true, terminal["CleanupInvoked"]);
        }

        [TestMethod]
        public void StableBaselineIsIndependentlyComparedBeforeTheFirstPreferenceWrite()
        {
            var probe = new Probe(); probe.Create(true).Run();
            Assert.IsTrue(probe.Phases.IndexOf("BaselineStabilityVerified") < probe.Phases.IndexOf("FontIntent"));
            Assert.AreEqual(probe.BaselineVersion, probe.Version());
            Assert.AreEqual(1, probe.Cleanup);
        }

        [TestMethod]
        public void InconsistentBaselineSnapshotFailsBeforeAnyPreferenceMutation()
        {
            var probe = new Probe { SemanticFault = "BaselineStabilityIntent" };
            Failure(() => probe.Create(true).Run());
            Assert.AreEqual(0, probe.Writes);
            Assert.IsFalse(probe.Phases.Contains("FontIntent"));
            Assert.AreEqual(1, probe.Cleanup, "A complete independently verified baseline restoration still permits normal cleanup.");
        }

        [DataTestMethod]
        [DataRow("Lost"), DataRow("Timeout"), DataRow("Null"), DataRow("Refused"), DataRow("MissingOk"), DataRow("NoData")]
        [DataRow("TopPending"), DataRow("Pending"), DataRow("Uncertain"), DataRow("DeliveryUncertain"), DataRow("VerificationPending"), DataRow("OpenDialog")]
        public void EveryNativeMatrixAndRestorationEmissionStopsAfterAnUnsafeReply(string fault)
        {
            // Obtain the complete predeclared successful sequence, including the expected refusals and restorations.
            var complete = new Probe { EmptySizes = true }; complete.Create().Run();
            for (int position = 1; position <= complete.Requests.Count; position++)
            {
                var probe = new Probe { EmptySizes = true, Fault = fault, FaultAt = position }; var runner = probe.Create();
                var failure = Failure(runner.Run);
                Assert.IsNotNull(failure, fault + " at " + position);
                Assert.AreEqual(position, probe.Requests.Count, "No observation, restoration or later scenario after " + fault + " at " + position);
                Assert.AreEqual(0, probe.Cleanup); Assert.AreEqual(1, probe.Preserved); Assert.IsTrue(runner.HostRetained);
                Assert.IsFalse(probe.Phases.Contains("CleanupIntent"));
                var retained = Copy(probe.Evidence.Last(item => item.Item1 == "HostRetained").Item2);
                Assert.AreEqual(false, retained["NativeReplayAllowed"]); Assert.AreEqual(false, retained["CleanupAllowed"]);
                Assert.IsTrue(retained.ContainsKey("Response"), "The original uncertain reply, including null, must be retained.");
                var terminal = Copy(probe.Evidence.Single(item => item.Item1 == "QualificationTerminal").Item2);
                Assert.AreEqual(false, terminal["Verified"]); Assert.AreEqual(false, terminal["CleanupInvoked"]);
                Assert.AreEqual(true, terminal["HostRetained"]);
            }
        }

        [DataTestMethod, DataRow("NoCommit"), DataRow("NoVerification")]
        public void EveryMutationRequiresExplicitCommittedVerifiedClosedResultBeforeRestorationLedger(string fault)
        {
            var complete = new Probe(); complete.Create().Run();
            for (int position = 1; position <= complete.Requests.Count; position++)
            {
                var request = complete.Requests[position - 1];
                if (!Equals(request["Command"], "set_vbe_option") || !request.ContainsKey("Query")) continue;
                var probe = new Probe { Fault = fault, FaultAt = position }; var runner = probe.Create(); Failure(runner.Run);
                Assert.AreEqual(position, probe.Requests.Count); Assert.AreEqual(0, probe.Cleanup); Assert.AreEqual(1, probe.Preserved);
                if (position == 3)
                {
                    var retained = Copy(probe.Evidence.Last(item => item.Item1 == "HostRetained").Item2);
                    Assert.AreEqual(0, ((System.Collections.ICollection)retained["CommittedRestoreEntries"]).Count, "An unqualified first write cannot enter restoration ledger.");
                }
            }
        }

        [DataTestMethod, DataRow("Throw"), DataRow("Null"), DataRow("ForeignPid"), DataRow("ProcessIdentityVerified"), DataRow("EnumerationSucceeded"), DataRow("OptionsDialogAbsent")]
        public void ExpectedRefusalCannotAuthorizeBridgeReadWithoutExactIndependentNativeClosure(string closureFault)
        {
            var probe = new Probe { EmptySizes = true, ClosureFault = closureFault }; var runner = probe.Create(); Failure(runner.Run);
            Assert.AreEqual(1, probe.ClosureCalls); Assert.AreEqual(0, probe.Cleanup); Assert.AreEqual(1, probe.Preserved);
            Assert.AreEqual("EmptySizeIntent", probe.Phases.Last(phase => phase.EndsWith("Intent")));
            Assert.IsFalse(probe.Phases.Contains("EmptySizeClosedReadbackIntent"));
        }

        [DataTestMethod, DataRow("WrongExpectedError"), DataRow("RefusalData"), DataRow("UnexpectedSuccess")]
        public void AnAmbiguousRefusalDoesNotEvenAuthorizeTheNativeClosureObservation(string fault)
        {
            var complete = new Probe { EmptySizes = true }; complete.Create().Run();
            int position = complete.Requests.FindIndex(item => item.TryGetValue("Property", out object property) && Equals(property, "Size")) + 1;
            var probe = new Probe { EmptySizes = true, FaultAt = position, Fault = fault }; var runner = probe.Create(); Failure(runner.Run);
            Assert.AreEqual(position, probe.Requests.Count); Assert.AreEqual(0, probe.ClosureCalls); Assert.AreEqual(0, probe.Cleanup);
            Assert.AreEqual(1, probe.Preserved);
        }

        [TestMethod]
        public void ExpectedRefusalStillFailsAndRetainsWhenIndependentClosedReadbackChangesRevision()
        {
            var probe = new Probe { EmptySizes = true, Fault = "RefusalChangedRevision" }; var runner = probe.Create();
            StringAssert.Contains(Failure(runner.Run).Message, "expected refusal changed");
            Assert.AreEqual(1, probe.ClosureCalls); Assert.AreEqual(0, probe.Cleanup); Assert.AreEqual(1, probe.Preserved);
            Assert.IsFalse(probe.Phases.Contains("BeforeRestorationIntent"));
        }

        [TestMethod]
        public void KnownCommittedWriteIsRestoredAfterIndependentSemanticAssertionFailure()
        {
            var probe = new Probe { SemanticFault = "FontIndependentAfterWriteIntent" }; var runner = probe.Create();
            StringAssert.Contains(Failure(runner.Run).Message, "native preference readback");
            Assert.AreEqual(2, probe.Writes); Assert.AreEqual(probe.BaselineVersion, probe.Version());
            Assert.AreEqual(1, probe.Cleanup); Assert.AreEqual(0, probe.Preserved);
            Assert.IsTrue(probe.Phases.Contains("BaselineRestored"));
        }

        [TestMethod]
        public void RestorationDoesNotEmitWhenFreshReadAlreadyMatchesBaseline()
        {
            var probe = new Probe { SemanticFault = "FontIndependentAfterWriteIntent", Fault = "AlreadyBaseline" };
            StringAssert.Contains(Failure(probe.Create().Run).Message, "native preference readback");
            Assert.AreEqual(1, probe.Writes); Assert.AreEqual(probe.BaselineVersion, probe.Version());
            Assert.IsTrue(probe.Phases.Contains("RestorationAlreadyMatched")); Assert.AreEqual(1, probe.Cleanup);
        }

        [TestMethod]
        public void ExactRestorationReadbackMismatchPreservesPrimaryAndStopsBeforeFurtherEntriesOrCleanup()
        {
            var probe = new Probe { SemanticFault = "FontIndependentAfterWriteIntent", Fault = "RestorationMismatch" };
            var error = Failure(probe.Create().Run) as AggregateException;
            Assert.IsNotNull(error); Assert.AreEqual(2, error.InnerExceptions.Count);
            StringAssert.Contains(error.InnerExceptions[0].Message, "native preference readback");
            StringAssert.Contains(error.InnerExceptions[1].Message, "Exact restoration readback");
            Assert.AreEqual(2, probe.Writes); Assert.AreEqual(0, probe.Cleanup); Assert.AreEqual(1, probe.Preserved);
            Assert.IsFalse(probe.Phases.Contains("CompleteRestorationReadbackIntent"));
        }

        [DataTestMethod, DataRow("FontIntent", 0), DataRow("FontReply", 2)]
        public void DurableEvidenceFailureBeforeOrAfterKnownCommitCannotInventOrLoseRestorationEntry(string phase, int writes)
        {
            var probe = new Probe { EvidenceFault = phase }; var runner = probe.Create();
            StringAssert.Contains(Failure(runner.Run).Message, "evidence failed");
            Assert.AreEqual(writes, probe.Writes); Assert.AreEqual(probe.BaselineVersion, probe.Version());
            Assert.AreEqual(1, probe.Cleanup); Assert.AreEqual(0, probe.Preserved);
        }

        [TestMethod]
        public void RefusalEvidenceFailureBeforeClosedProofStopsAllBridgeRequestsAndCleanup()
        {
            var probe = new Probe { EmptySizes = true, EvidenceFault = "EmptySizeRefusalReply" }; var runner = probe.Create();
            StringAssert.Contains(Failure(runner.Run).Message, "evidence failed");
            Assert.AreEqual(0, probe.ClosureCalls); Assert.AreEqual(0, probe.Cleanup); Assert.AreEqual(1, probe.Preserved);
            Assert.IsFalse(probe.Phases.Contains("EmptySizeClosedReadbackIntent"));
        }

        [DataTestMethod, DataRow("ScenarioMatrix"), DataRow("BaselineIntent"), DataRow("BaselineReply")]
        public void PreMutationEvidenceFailureHasNoRestoreWriteAndAllowsOneKnownSafeCleanup(string phase)
        {
            var probe = new Probe { EvidenceFault = phase }; StringAssert.Contains(Failure(probe.Create().Run).Message, "evidence failed");
            Assert.AreEqual(0, probe.Writes); Assert.AreEqual(1, probe.Cleanup); Assert.AreEqual(0, probe.Preserved);
        }

        [DataTestMethod, DataRow("EmptySizeNativeClosureObservation"), DataRow("RestorationReply"), DataRow("CleanupIntent")]
        public void FailureToRecordClosureOrRestorationOrCleanupIntentRetainsExactHostWithoutFurtherDispatch(string phase)
        {
            var probe = new Probe { EmptySizes = true, EvidenceFault = phase }; var runner = probe.Create();
            StringAssert.Contains(Failure(runner.Run).Message, "evidence failed");
            Assert.IsTrue(runner.HostRetained); Assert.AreEqual(0, probe.Cleanup); Assert.AreEqual(1, probe.Preserved);
            if (phase == "EmptySizeNativeClosureObservation") Assert.IsFalse(probe.Phases.Contains("EmptySizeClosedReadbackIntent"));
            if (phase == "RestorationReply") Assert.IsFalse(probe.Phases.Contains("RestorationReadbackIntent"));
        }

        [TestMethod]
        public void KnownPrimaryAndUncertainRestorationRemainDistinctAndNoLaterCompensationOccurs()
        {
            // Baseline, Font-before, Font-write, Font-independent, then the first restoration read.
            var probe = new Probe { SemanticFault = "FontIndependentAfterWriteIntent", FaultAt = 5, Fault = "Timeout" };
            var failure = Failure(probe.Create().Run) as AggregateException;
            Assert.IsNotNull(failure); Assert.AreEqual(2, failure.InnerExceptions.Count);
            StringAssert.Contains(failure.InnerExceptions[0].Message, "native preference readback");
            StringAssert.Contains(failure.InnerExceptions[1].Message, "timed out");
            Assert.AreEqual(5, probe.Requests.Count); Assert.AreEqual(1, probe.Writes); Assert.AreEqual(0, probe.Cleanup);
        }

        [TestMethod]
        public void IncompleteBaselineRestorationNeverPermitsCleanupOrClaimsQualification()
        {
            var probe = new Probe { Fault = "BaselineMismatch" }; var runner = probe.Create();
            StringAssert.Contains(Failure(runner.Run).Message, "Every category and preference");
            Assert.AreEqual(0, probe.Cleanup); Assert.AreEqual(1, probe.Preserved); Assert.IsFalse(probe.Phases.Contains("BaselineRestored"));
        }

        [TestMethod]
        public void PrimaryAndCleanupFailuresAndUnknownReplyPreservationErrorsRemainDistinct()
        {
            var known = new Probe { SemanticFault = "FontIndependentAfterWriteIntent", CleanupFault = true };
            var failure = Failure(known.Create().Run) as AggregateException;
            Assert.IsNotNull(failure); Assert.AreEqual(2, failure.InnerExceptions.Count);
            StringAssert.Contains(failure.InnerExceptions[0].Message, "native preference readback");
            StringAssert.Contains(failure.InnerExceptions[1].Message, "cleanup failed");
            var terminal = Copy(known.Evidence.Single(item => item.Item1 == "QualificationTerminal").Item2);
            StringAssert.Contains((string)terminal["PrimaryError"], "native preference readback");
            StringAssert.Contains((string)terminal["ShutdownError"], "cleanup failed");
            var unknown = new Probe { FaultAt = 3, Fault = "Lost", PreservationFault = true, EvidenceFault = "HostRetained" };
            var retained = Failure(unknown.Create().Run) as AggregateException;
            Assert.IsNotNull(retained); Assert.AreEqual(3, retained.InnerExceptions.Count);
            StringAssert.Contains(retained.InnerExceptions[0].Message, "lost"); StringAssert.Contains(retained.InnerExceptions[1].Message, "retention failed");
            StringAssert.Contains(retained.InnerExceptions[2].Message, "evidence failed");
            Assert.AreEqual(3, unknown.Requests.Count); Assert.AreEqual(0, unknown.Cleanup);
        }

        [TestMethod]
        public void TerminalEvidenceFailureAfterVerifiedRestorationAndShutdownRemainsFailureWithoutCleanupReplay()
        {
            var probe = new Probe { EvidenceFault = "QualificationTerminal" };
            StringAssert.Contains(Failure(probe.Create().Run).Message, "evidence failed");
            Assert.AreEqual(1, probe.Cleanup); Assert.AreEqual(0, probe.Preserved); Assert.AreEqual(probe.BaselineVersion, probe.Version());
        }

        [TestMethod]
        public void EveryReadRequiresAnExactFullRevisionEvenWhenDialogIsClosed()
        {
            var complete = new Probe(); complete.Create().Run();
            for (int position = 1; position <= complete.Requests.Count; position++)
            {
                if (!Equals(complete.Requests[position - 1]["Command"], "read_vbe_options")) continue;
                var probe = new Probe { Fault = "BadRevision", FaultAt = position }; var runner = probe.Create();
                StringAssert.Contains(Failure(runner.Run).Message, "SHA-256");
                Assert.AreEqual(position, probe.Requests.Count); Assert.AreEqual(0, probe.Cleanup); Assert.AreEqual(1, probe.Preserved);
            }
        }
    }
}
