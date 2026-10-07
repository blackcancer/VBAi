using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.ExceptionServices;
using System.Text.RegularExpressions;

namespace VBAi.Tests.Integration
{
    /// <summary>Test-local full Format matrix. Unknown native outcomes suspend all dispatch and cleanup.</summary>
    internal sealed class ExcelFormatOptionsQualification
    {
        internal static readonly string[] Scenarios = { "exact font", "size catalogue or honest refusal", "foreground", "background", "indicator",
            "nondefault category foreground", "margin indicator", "stale version refuses", "full options version restored" };
        private readonly Func<object, IDictionary<string, object>> dispatch;
        private readonly Func<IDictionary<string, object>> observeClosure;
        private readonly Action preserve, cleanup;
        private readonly Action<string, object> evidence;
        private readonly int processId;
        private readonly bool verifyReadStability, marginOnly, historicalPalettePrefix, fontSizeOnly, historicalFullMatrix;
        private readonly List<Tuple<string, string, object, string>> ledger = new List<Tuple<string, string, object, string>>();
        private IDictionary<string, object> baseline;
        private readonly Action verifyExclusiveHost;
        internal bool HostRetained { get; private set; }

        /// <summary>Compile the unchanged native guard through one known refusal, before attaching its exact IL breakpoint.</summary>
        internal void WarmGuardForBreakpoint()
        {
            var before = Read("GuardBreakpointWarmupBefore");
            Refusal("GuardBreakpointWarmup", new
            {
                Command = "set_vbe_option",
                Pane = (string)Format(before)["Tab"],
                Property = "__Q026_READ_ONLY_GUARD_WARMUP__",
                Value = "No preference mutation",
                ExpectedOptionsVersion = new string('0', 64)
            }, before, "VBE options changed since inspection; read them again.");
            evidence("GuardBreakpointWarmupVerified", new { NativePreferenceWrites = 0, FailedMutationReplayed = false });
        }

        internal ExcelFormatOptionsQualification(int processId, Func<object, IDictionary<string, object>> dispatch,
            Func<IDictionary<string, object>> observeClosure, Action preserve, Action cleanup, Action<string, object> evidence,
            bool verifyReadStability = false, bool marginOnly = false, bool historicalPalettePrefix = false,
            Action verifyExclusiveHost = null, bool fontSizeOnly = false, bool historicalFullMatrix = false)
        {
            this.processId = processId; this.dispatch = dispatch; this.observeClosure = observeClosure;
            this.preserve = preserve; this.cleanup = cleanup; this.evidence = evidence; this.verifyReadStability = verifyReadStability; this.marginOnly = marginOnly;
            this.historicalPalettePrefix = historicalPalettePrefix; this.verifyExclusiveHost = verifyExclusiveHost;
            this.fontSizeOnly = fontSizeOnly; this.historicalFullMatrix = historicalFullMatrix;
            if (historicalFullMatrix && (!historicalPalettePrefix || marginOnly || fontSizeOnly))
                throw new ArgumentException("Historical full matrix requires historical mutation order and the complete scope.");
        }

        /// <summary>Validate the Format opt-in before preparation, then hand off only a successfully owned bootstrap.</summary>
        internal static void RunOwned<T>(bool enabled, string ownedResults, string evidenceRoot, string inheritedDiagnosticManifest,
            Action prepareEvidence, Func<string, T> bootstrap, Action<T> qualify) where T : class
        {
            if (!enabled) Assert.Inconclusive("Excel automation is opt-in. Set VBAi_RUN_EXCEL_TESTS=1.");
            ExcelOwnedBootstrapPlan.RequireLocalAbsolutePath(ownedResults);
            string exactEvidence = ExcelOwnedBootstrapPlan.RequireLocalAbsolutePath(evidenceRoot);
            if (!string.IsNullOrEmpty(inheritedDiagnosticManifest))
                throw new InvalidOperationException("Format qualification must not inherit a path-visibility/token manifest; no host was launched.");
            prepareEvidence();
            // StartOwnedWithTrace owns failure retention. Never wrap startup in a Close/Quit finally or activation fallback.
            T host = bootstrap(Path.Combine(exactEvidence, "owned-bootstrap-phases.jsonl"));
            if (host == null) throw new InvalidOperationException("The owned bootstrap returned no ready fixture; no qualification dispatch or cleanup is permitted.");
            qualify(host);
        }

        internal void Run()
        {
            Exception primary = null, restoration = null, shutdown = null, terminalRecord = null;
            bool cleanupInvoked = false;
            try { Matrix(); } catch (Exception error) { primary = error; }
            if (!HostRetained && baseline != null)
                try { Restore(); } catch (Exception error) { restoration = HostRetained ? error : Retain(error); }
            if (!HostRetained)
            {
                try { evidence("CleanupIntent", new { NativeRequestsComplete = true }); cleanupInvoked = true; cleanup(); }
                catch (Exception error) { shutdown = cleanupInvoked ? error : Retain(error); }
            }
            try
            {
                evidence("QualificationTerminal", new
                {
                    Verified = primary == null && restoration == null && shutdown == null,
                    HostRetained,
                    CleanupInvoked = cleanupInvoked,
                    PrimaryError = primary?.ToString(),
                    RestorationError = restoration?.ToString(),
                    ShutdownError = shutdown?.ToString(),
                    NativeReplayAllowed = false
                });
            }
            catch (Exception error) { terminalRecord = error; }
            ThrowFailures("Format scenario, restoration, cleanup and terminal evidence failures are preserved separately.", primary, restoration, shutdown, terminalRecord);
        }

        private void Matrix()
        {
            evidence("ScenarioMatrix", marginOnly ? new[] { "margin indicator", "full options version restored" } :
                fontSizeOnly ? new[] { "exact font", "size catalogue or honest refusal", "full options version restored" } :
                historicalPalettePrefix && !historicalFullMatrix ? new[] { "historical font and size refusal", "historical foreground/background/indicator without Query", "full options version restored" } : Scenarios);
            baseline = Read("Baseline");
            evidence("BaselineComplete", baseline);
            if (verifyReadStability)
            {
                var stable = Read("BaselineStability");
                Assert.AreEqual(Version(baseline), Version(stable), "Complete revision drifted before any preference write.");
                var json = new System.Web.Script.Serialization.JavaScriptSerializer();
                Assert.AreEqual(json.Serialize(baseline["Tabs"]), json.Serialize(stable["Tabs"]),
                    "Hash equality must not hide a different recorded options snapshot.");
                evidence("BaselineStabilityVerified", new { Before = baseline, After = stable, PreferenceWrites = 0 });
            }
            var format = Format(baseline); string tab = (string)format["Tab"];
            if (marginOnly) { Margin(tab); return; }
            var font = Find(format, "Font", "Police :");
            string alternate = Choices(font).FirstOrDefault(x => x != (string)font["Value"] &&
                new[] { "Consolas (Occidental)", "Consolas (Western)", "Consolas", "Courier New (Occidental)", "Courier New (Western)", "Courier New" }.Contains(x));
            Assert.IsNotNull(alternate, "An exact observed alternate font is required for this bounded fixture.");
            Write("Font", tab, (string)font["Name"], alternate, null);
            Assert.AreEqual(alternate, Find(Format(Read("FontReadback")), "Font", "Police :")["Value"]);

            var sizeRead = Read("SizeCatalogue"); var size = Find(Format(sizeRead), "Size", "Taille :");
            var sizes = Choices(size);
            if (sizes.Length == 0)
            {
                Assert.IsNotNull(size["Value"], "Empty native catalogue must still report the real edit value.");
                Refusal("EmptySize", new
                {
                    Command = "set_vbe_option",
                    Pane = tab,
                    Property = size["Name"],
                    Value = "12",
                    ExpectedOptionsVersion = Version(sizeRead)
                }, sizeRead, "The exact native choice is absent, ambiguous or unreadable.");
            }
            else
            {
                string next = sizes.First(x => x != (string)size["Value"]);
                Write("Size", tab, (string)size["Name"], next, null);
                Assert.AreEqual(next, Find(Format(Read("SizeReadback")), "Size", "Taille :")["Value"]);
            }
            if (fontSizeOnly) return; // Keep the catalogue diagnostic inside the same restoration and owned shutdown lifecycle.
            foreach (var names in new[] { new[] { "Foreground", "Premier plan :" }, new[] { "Background", "Arrière-plan :" }, new[] { "Indicator", "Indicateur :" } })
            {
                var paletteFormat = Format(Read(names[0] + "Catalogue")); string category = CurrentCategory(paletteFormat);
                var control = historicalPalettePrefix ? Find(paletteFormat, names) : CategoryPalette(paletteFormat, category, names);
                string next = Choices(control).First(x => x != (string)control["Value"]);
                Write(names[0], tab, (string)control["Name"], next, historicalPalettePrefix ? null : category);
                var readback = Format(Read(names[0] + "Readback"));
                Assert.AreEqual(next, (historicalPalettePrefix ? Find(readback, names) : CategoryPalette(readback, category, names))["Value"]);
            }
            // The original rethrow stack does not identify the failing stage. The
            // prefix deliberately omits later stages; a pass cannot exclude them.
            if (historicalPalettePrefix && !historicalFullMatrix) return;
            var categoryRead = Read("OtherCategoryCatalogue"); var categories = (object[])Format(categoryRead)["FormatCategories"];
            Assert.IsTrue(categories.Length > 1);
            var other = VbeBridgeClient.Object(categories[1]);
            var foreground = ((object[])other["Palettes"]).Select(VbeBridgeClient.Object)
                .Single(x => new[] { "Foreground", "Premier plan :" }.Contains((string)x["Name"]));
            string otherCategory = (string)other["Category"];
            string colour = Choices(foreground).First(x => x != (string)foreground["Value"]);
            Write("OtherCategory", tab, (string)foreground["Name"], colour, otherCategory);
            Assert.AreEqual(colour, CategoryPalette(Format(Read("OtherCategoryReadback")), otherCategory, (string)foreground["Name"])["Value"]);
            Margin(tab);
            var beforeStale = Read("BeforeStaleRefusal");
            Assert.AreNotEqual(Version(baseline), Version(beforeStale), "The stale scenario needs a genuinely changed complete revision.");
            Refusal("StaleVersion", new
            {
                Command = "set_vbe_option",
                Pane = tab,
                Property = font["Name"],
                Value = font["Value"],
                ExpectedOptionsVersion = Version(baseline)
            }, beforeStale, "VBE options changed since inspection; read them again.");
        }

        private void Margin(string tab)
        {
            var margin = Find(Format(Read("MarginCatalogue")), "Margin Indicator Bar", "Barre des indicateurs en marge");
            Assert.IsTrue(Equals(margin["Value"], "On") || Equals(margin["Value"], "Off"), "The native margin value must be exact On/Off.");
            bool nextMargin = !Equals(margin["Value"], "On");
            Write("Margin", tab, (string)margin["Name"], nextMargin, null);
            Assert.AreEqual(nextMargin ? "On" : "Off", Find(Format(Read("MarginReadback")), (string)margin["Name"])["Value"]);
        }

        private void Restore()
        {
            // Only positively committed, closed replies entered this ledger. Stop on the first unsafe restoration.
            foreach (var entry in ledger.AsEnumerable().Reverse())
            {
                var current = Read("BeforeRestoration");
                object actual = Control(Format(current), entry.Item2, entry.Item4)["Value"];
                if (Equals(actual, Expected(entry.Item3)))
                { evidence("RestorationAlreadyMatched", new { entry.Item2, entry.Item4, Expected = entry.Item3, Before = current }); continue; }
                Send("Restoration", new
                {
                    Command = "set_vbe_option",
                    Pane = entry.Item1,
                    Property = entry.Item2,
                    Value = entry.Item3,
                    Query = entry.Item4,
                    ExpectedOptionsVersion = Version(current)
                }, true, null);
                var after = Read("RestorationReadback");
                Assert.AreEqual(Expected(entry.Item3), Control(Format(after), entry.Item2, entry.Item4)["Value"], "Exact restoration readback is required.");
                evidence("RestorationEntryVerified", new { Property = entry.Item2, Category = entry.Item4, Readback = after });
            }
            var restored = Read("CompleteRestorationReadback");
            Assert.AreEqual(Version(baseline), Version(restored), "Every category and preference must return to the complete baseline revision.");
            var json = new System.Web.Script.Serialization.JavaScriptSerializer();
            Assert.AreEqual(json.Serialize(baseline["Tabs"]), json.Serialize(restored["Tabs"]),
                "Every recorded tab, catalogue and palette must match the complete baseline, independently of its hash.");
            evidence("BaselineRestored", new { BaselineVersion = Version(baseline), Readback = restored });
        }

        private void Write(string phase, string tab, string property, object value, string category)
        {
            var before = Read(phase + "BeforeWrite");
            object old = Control(Format(before), property, category)["Value"];
            if (value is bool) old = Equals(old, "On");
            // Keep the historical request's null Query, but bind its positive
            // recovery entry to the actual category. Later Query changes otherwise
            // redirect an implicit palette compensation to a different category.
            string restoreCategory = historicalFullMatrix && category == null &&
                new[] { "Foreground", "Background", "Indicator" }.Contains(phase)
                ? CurrentCategory(Format(before)) : category;
            Send(phase, new
            {
                Command = "set_vbe_option",
                Pane = tab,
                Property = property,
                Value = value,
                Query = category,
                ExpectedOptionsVersion = Version(before)
            }, true, () =>
            {
                if (!ledger.Any(item => item.Item2 == property && item.Item4 == restoreCategory))
                    ledger.Add(Tuple.Create(tab, property, old, restoreCategory));
            });
            if (historicalPalettePrefix)
            {
                // The original test reads back in Matrix immediately after Write. Keep that
                // order rather than inserting another full native inspection between them.
                evidence(phase + "HistoricalCommit", new
                {
                    Before = before,
                    MutationRetried = false,
                    IndependentReadbackRequiredByMatrix = true
                });
                return;
            }
            var after = Read(phase + "IndependentAfterWrite");
            Assert.AreEqual(Expected(value), Control(Format(after), property, category)["Value"], "The native preference readback differs from the requested value.");
            evidence(phase + "Verified", new
            {
                Before = before,
                IndependentAfterRead = after,
                ObservedCategoryBefore = CurrentCategory(Format(before)),
                ObservedCategoryAfter = CurrentCategory(Format(after)),
                MutationRetried = false
            });
        }

        private void Refusal(string phase, object request, IDictionary<string, object> before, string expectedError)
        {
            EnsureDispatchAllowed(); evidence(phase + "Intent", request);
            IDictionary<string, object> reply;
            try { reply = dispatch(request); } catch (Exception error) { throw OutcomeFailure(error, request, null); }
            if (reply == null || !reply.TryGetValue("Ok", out object ok) || !Equals(ok, false) || Unsafe(reply) ||
                !reply.TryGetValue("Error", out object errorText) || !Equals(errorText, expectedError) ||
                (reply.TryGetValue("Data", out object raw) && raw != null))
                throw OutcomeFailure(new InvalidOperationException("The anticipated native refusal is absent or ambiguous; no further dispatch is permitted."), request, reply);
            // Error wording does not establish that the Options dialog actually closed. No bridge request is allowed yet.
            try
            {
                evidence(phase + "RefusalReply", reply);
                var observation = observeClosure();
                evidence(phase + "NativeClosureObservation", observation);
                if (observation == null || !observation.TryGetValue("ProcessId", out object pid) || !Equals(pid, processId) ||
                    !True(observation, "ProcessIdentityVerified") || !True(observation, "EnumerationSucceeded") || !True(observation, "OptionsDialogAbsent"))
                    throw new InvalidOperationException("Native Options-window absence for the exact owned process was not independently verified.");
            }
            catch (Exception error) { throw OutcomeFailure(error, request, reply); }
            var after = Read(phase + "ClosedReadback");
            if (Version(before) != Version(after))
                throw Retain(new InvalidOperationException("The expected refusal changed the complete preferences revision; do not continue or replay."));
            var serializer = new System.Web.Script.Serialization.JavaScriptSerializer();
            if (serializer.Serialize(before["Tabs"]) != serializer.Serialize(after["Tabs"]))
                throw Retain(new InvalidOperationException("The expected refusal changed the complete preferences structure; do not continue or replay."));
            evidence(phase + "VerifiedRefusal", new { Response = reply, IndependentClosedReadback = after, MutationRetried = false });
        }

        private IDictionary<string, object> Read(string phase) => Send(phase, new { Command = "read_vbe_options" }, false, null);
        private IDictionary<string, object> Send(string phase, object request, bool mutation, Action terminal)
        {
            EnsureDispatchAllowed(); evidence(phase + "Intent", request);
            IDictionary<string, object> reply;
            try { reply = dispatch(request); } catch (Exception error) { throw OutcomeFailure(error, request, null); }
            if (reply == null || !reply.TryGetValue("Ok", out object ok) || !Equals(ok, true) || Unsafe(reply) ||
                !reply.TryGetValue("Data", out object raw) || !(raw is IDictionary<string, object> data) || Unsafe(data) ||
                !True(data, "DialogClosed") || (mutation && (!True(data, "CommitRequested") || !True(data, "ControlValueVerified"))))
                throw OutcomeFailure(new InvalidOperationException("The Format reply is lost, refused, pending or incomplete; preserve without further native dispatch or cleanup."), request, reply);
            if (!mutation)
                try { Version(data); } catch (Exception error) { throw OutcomeFailure(error, request, reply); }
            terminal?.Invoke(); // A known native commit is recorded even if durable evidence subsequently fails.
            evidence(phase + "Reply", new { Response = reply, MutationRetried = false });
            return data;
        }
        private void EnsureDispatchAllowed()
        {
            if (HostRetained) throw new InvalidOperationException("The exact owned host is retained; dispatch and cleanup are forbidden.");
            try { verifyExclusiveHost?.Invoke(); }
            catch (Exception error) { throw Retain(error); }
        }
        private Exception OutcomeFailure(Exception error, object request, object response)
        { return Retain(error, request, response); }
        private Exception Retain(Exception primary, object request = null, object response = null)
        {
            Exception retention = null, recording = null;
            if (!HostRetained) { HostRetained = true; try { preserve(); } catch (Exception error) { retention = error; } }
            try
            {
                evidence("HostRetained", new
                {
                    Error = primary.ToString(),
                    Request = request,
                    Response = response,
                    NativeReplayAllowed = false,
                    CleanupAllowed = false,
                    CommittedRestoreEntries = ledger.Select(item => new { Pane = item.Item1, Property = item.Item2, Value = item.Item3, Category = item.Item4 }).ToArray()
                });
            }
            catch (Exception error) { recording = error; }
            var failures = new[] { primary, retention, recording }.Where(error => error != null).ToArray();
            return failures.Length == 1 ? primary : new AggregateException("Native failure and retention/evidence failures are preserved.", failures);
        }
        private static bool Unsafe(IDictionary<string, object> data) => new[] { "Pending", "Uncertain", "DeliveryUncertain", "VerificationPending" }
            .Any(key => data.TryGetValue(key, out object value) && !Equals(value, false) && value != null);
        private static bool True(IDictionary<string, object> data, string key) => data.TryGetValue(key, out object value) && Equals(value, true);
        private static string Version(IDictionary<string, object> data)
        {
            string value = data.TryGetValue("OptionsVersion", out object raw) ? raw as string : null;
            if (value == null || !Regex.IsMatch(value, "^[a-fA-F0-9]{64}$")) throw new InvalidOperationException("The complete options SHA-256 revision is absent or malformed."); return value;
        }
        private static object Expected(object value) => value is bool check ? (object)(check ? "On" : "Off") : value;
        private static IDictionary<string, object> Format(IDictionary<string, object> data) => ((object[])data["Tabs"]).Select(VbeBridgeClient.Object)
            .Single(item => new[] { "Editor Format", "Format de l'éditeur", "Format de l’éditeur" }.Contains((string)item["Tab"]));
        private static IDictionary<string, object> Find(IDictionary<string, object> format, params string[] names) => ((object[])format["Controls"]).Select(VbeBridgeClient.Object)
            .Single(item => names.Contains((string)item["Name"]) && (string)item["Type"] != "ControlType.Text");
        private static IDictionary<string, object> Control(IDictionary<string, object> format, string property, string category) =>
            string.IsNullOrEmpty(category) ? Find(format, property) : CategoryPalette(format, category, property);
        private static IDictionary<string, object> CategoryPalette(IDictionary<string, object> format, string category, params string[] names) =>
            ((object[])format["FormatCategories"]).Select(VbeBridgeClient.Object).Single(item => Equals(item["Category"], category))["Palettes"] is object[] palettes
                ? palettes.Select(VbeBridgeClient.Object).Single(item => names.Contains((string)item["Name"])) : throw new InvalidOperationException("Native palettes are absent.");
        private static string[] Choices(IDictionary<string, object> control) => ((object[])control["Choices"]).Cast<string>().ToArray();
        private static string CurrentCategory(IDictionary<string, object> format)
        {
            var control = ((object[])format["Controls"]).Select(VbeBridgeClient.Object).Single(item => (string)item["Type"] == "ControlType.List" &&
                new[] { "code colors", "couleurs du code", "color text", "texte couleur" }.Contains(((string)item["Name"] ?? "").Replace("&", "").Trim().TrimEnd(':').Trim().ToLowerInvariant()));
            string category = control["Value"] as string;
            Assert.IsFalse(string.IsNullOrWhiteSpace(category)); Assert.AreEqual(1, Choices(control).Count(item => item == category)); return category;
        }
        private static void ThrowFailures(string message, params Exception[] failures)
        {
            var actual = failures.Where(error => error != null).ToArray(); if (actual.Length > 1) throw new AggregateException(message, actual);
            if (actual.Length == 1) ExceptionDispatchInfo.Capture(actual[0]).Throw();
        }
    }
}
