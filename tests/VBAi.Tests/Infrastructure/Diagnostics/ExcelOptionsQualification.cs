using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.ExceptionServices;
using System.Text.RegularExpressions;

namespace VBAi.Tests.Integration
{
    /// <summary>Test-local options lifecycle; uncertain delivery never permits another native request or cleanup.</summary>
    internal sealed class ExcelOptionsQualification
    {
        private readonly Func<object, IDictionary<string, object>> dispatch;
        private readonly Action preserve, cleanup;
        private readonly Action<string, object> evidence;
        internal bool HostRetained { get; private set; }

        internal ExcelOptionsQualification(Func<object, IDictionary<string, object>> dispatch,
            Action preserve, Action cleanup, Action<string, object> evidence)
        { this.dispatch = dispatch; this.preserve = preserve; this.cleanup = cleanup; this.evidence = evidence; }

        internal void Run()
        {
            Exception primary = null, shutdown = null;
            try { RoundTrip(); }
            catch (Exception error) { primary = error; }
            if (!HostRetained)
            {
                try { evidence("CleanupIntent", new { NativeRequestsComplete = true }); cleanup(); }
                catch (Exception error) { shutdown = error; }
            }
            ThrowFailures("Options scenario and cleanup failed; both errors are retained.", primary, shutdown);
        }

        private void RoundTrip()
        {
            var baseline = Read("Baseline");
            string baselineVersion = Version(baseline);
            var tabs = Items(baseline["Tabs"]);
            foreach (string fragment in new[] { "format", "ancrage" })
            {
                var tab = tabs.Single(item => Text(item["Tab"]).ToLowerInvariant().Contains(fragment));
                var checkbox = Items(tab["Controls"]).First(item => Text(item["Type"]) == "ControlType.CheckBox" && item["Error"] == null);
                string originalValue = Text(checkbox["Value"]);
                if (originalValue != "On" && originalValue != "Off")
                    throw new InvalidOperationException("The baseline checkbox is not an exact On/Off value; no write is permitted.");
                bool original = originalValue == "On";
                string pane = Text(tab["Tab"]), property = Text(checkbox["Name"]);
                bool knownWrite = false;
                Exception primary = null, restoration = null;
                try
                {
                    var before = Read("BeforeMutation");
                    Write("Mutation", pane, property, !original, Version(before), () => knownWrite = true);
                    var after = Read("MutationReadback");
                    var observedTab = Items(after["Tabs"]).Single(item => Text(item["Tab"]) == pane);
                    var observed = Items(observedTab["Controls"]).Single(item => Text(item["Name"]) == property && Text(item["Type"]) == Text(checkbox["Type"]));
                    if (Text(observed["Value"]) != (original ? "Off" : "On"))
                        throw new InvalidOperationException("The native checkbox readback differs from the requested value.");
                }
                catch (Exception error) { primary = error; }
                finally
                {
                    // A pre-write failure cannot authorize a preference write. A lost reply forbids all further dispatch.
                    if (knownWrite && !HostRetained)
                    {
                        try
                        {
                            var beforeRestore = Read("BeforeRestoration");
                            Write("Restoration", pane, property, original, Version(beforeRestore), () => { });
                            var restored = Read("RestorationReadback");
                            if (Version(restored) != baselineVersion)
                                throw new InvalidOperationException("The complete native preferences were not restored to the baseline revision.");
                            evidence("BaselineRestored", new { OptionsVersion = baselineVersion });
                        }
                        catch (Exception error) { restoration = Retain(error); }
                    }
                    else
                    {
                        try { evidence("RestorationNotEmitted", new { KnownTerminalWrite = knownWrite, HostRetained }); }
                        catch (Exception error) { restoration = error; }
                    }
                }
                ThrowFailures("Options mutation and restoration failed; both errors are retained.", primary, restoration);
            }
        }

        private IDictionary<string, object> Read(string phase)
        {
            var data = Send(phase, new { Command = "read_vbe_options" }, false, null);
            evidence(phase + "Summary", Summary(data));
            return data;
        }

        private void Write(string phase, string pane, string property, bool value, string version, Action terminal)
        {
            Send(phase, new { Command = "set_vbe_option", Pane = pane, Property = property,
                Value = value, ExpectedOptionsVersion = version }, true, terminal);
        }

        private IDictionary<string, object> Send(string phase, object request, bool mutation, Action terminal)
        {
            if (HostRetained) throw new InvalidOperationException("The host is retained; no further native request is permitted.");
            evidence(phase + "Intent", request); // Durable before the first possible emission.
            IDictionary<string, object> reply;
            try { reply = dispatch(request); }
            catch (Exception error) { throw Retain(error); }
            if (reply == null || !reply.TryGetValue("Ok", out object ok) || !Equals(ok, true) ||
                !reply.TryGetValue("Data", out object raw) || !(raw is IDictionary<string, object> data))
                throw Retain(new InvalidOperationException("The options reply is absent, refused or incomplete; native outcome is not qualified."));
            if (True(data, "Pending") || True(data, "Uncertain") || True(data, "DeliveryUncertain") ||
                !True(data, "DialogClosed") || (mutation && (!True(data, "CommitRequested") || !True(data, "ControlValueVerified"))))
                throw Retain(new InvalidOperationException("The options operation is not known terminal and closed; preserve without native replay."));
            terminal?.Invoke(); // Record known completion even if subsequent evidence writing fails.
            evidence(phase + "Reply", mutation ? (object)data : new { OptionsVersion = Version(data), DialogClosed = true });
            return data;
        }

        private Exception Retain(Exception primary)
        {
            Exception retention = null, recording = null;
            if (!HostRetained)
            {
                HostRetained = true;
                try { preserve(); } catch (Exception error) { retention = error; }
            }
            string text = primary.ToString();
            try { evidence("HostRetained", new { Error = Bound(text, 8192), OriginalErrorCharacters = text.Length,
                ErrorTruncated = text.Length > 8192, NativeReplayAllowed = false, CleanupAllowed = false }); }
            catch (Exception error) { recording = error; }
            var failures = new[] { primary, retention, recording }.Where(error => error != null).ToArray();
            return failures.Length == 1 ? primary : new AggregateException("Options failure and preservation/evidence errors are retained.", failures);
        }

        internal static object Summary(IDictionary<string, object> data)
        {
            var tabs = Items(data["Tabs"]);
            return new { OptionsVersion = Version(data), DialogClosed = data["DialogClosed"],
                TabCount = tabs.Length, Tabs = tabs.Take(8).Select(tab => new {
                    Tab = Bound(Text(tab["Tab"]), 256), ControlCount = Items(tab["Controls"]).Length,
                    Controls = Items(tab["Controls"]).Take(128).Select(control => new {
                        Name = Bound(Text(control["Name"]), 256), Type = Bound(Text(control["Type"]), 64),
                        Value = Bound(Text(control["Value"]), 128), ValueTruncated = Text(control["Value"]).Length > 128,
                        Error = Bound(Text(control["Error"]), 256), ErrorTruncated = Text(control["Error"]).Length > 256 }).ToArray(),
                    ControlsOmitted = Math.Max(0, Items(tab["Controls"]).Length - 128) }).ToArray(),
                TabsOmitted = Math.Max(0, tabs.Length - 8), CataloguesOmitted = true };
        }

        private static string Version(IDictionary<string, object> data)
        {
            string version = data.TryGetValue("OptionsVersion", out object value) ? Text(value) : null;
            if (version == null || !Regex.IsMatch(version, "^[a-fA-F0-9]{64}$"))
                throw new InvalidOperationException("The complete options SHA-256 revision is absent or malformed.");
            return version;
        }
        private static IDictionary<string, object>[] Items(object values) => ((object[])values).Cast<IDictionary<string, object>>().ToArray();
        private static bool True(IDictionary<string, object> data, string key) => data.TryGetValue(key, out object value) && Equals(value, true);
        private static string Text(object value) => Convert.ToString(value);
        private static string Bound(string value, int limit) => value.Length <= limit ? value : value.Substring(0, limit);
        private static void ThrowFailures(string message, params Exception[] failures)
        {
            var actual = failures.Where(error => error != null).ToArray();
            if (actual.Length > 1) throw new AggregateException(message, actual);
            if (actual.Length == 1) ExceptionDispatchInfo.Capture(actual[0]).Throw();
        }
    }
}
