using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading;

namespace VBAi.Tests.Integration
{
    /// <summary>Adapts existing owned-host fixtures without bypassing their containment or shutdown guards.</summary>
    internal sealed class OllamaOfficeHost : IDisposable
    {
        internal delegate IDictionary<string, object> ReadData(string command, params object[] pairs);
        internal ReadData Data;
        internal Func<string, IDictionary<string, object>[]> Items;
        internal string Project, Root;
        private string projectName, projectPath;
        internal string ToolProject => !string.IsNullOrWhiteSpace(projectPath) && System.IO.Path.IsPathRooted(projectPath) ? projectPath : projectName;
        internal string Label => projectName + " · " + (!string.IsNullOrWhiteSpace(projectPath) &&
            System.IO.Path.IsPathRooted(projectPath) ? System.IO.Path.GetFileName(projectPath) : UiText.Get("unsaved document"));
        internal int ProcessId;
        internal Action Retain;
        private IDisposable fixture;
        private bool retained;
        internal bool NativeDispatchUnsettled { get; private set; }
        internal bool IsRetained => retained;
        private static readonly List<OllamaOfficeHost> Retained = new List<OllamaOfficeHost>();

        internal static OllamaOfficeHost Start(string kind, Action<OllamaOfficeHost> assigned)
        {
            if (OllamaOfficeDesktop.MainEnabled) OllamaOfficeDesktop.Require(null);
            var result = new OllamaOfficeHost();
            Action<OllamaOfficeHost> publish = value =>
            {
                var retainFixture = value.Retain;
                value.Retain = () => { if (value.retained) return; value.retained = true; Retained.Add(value); retainFixture(); };
                assigned(value); // Transfer exact fixture ownership before any structural validation.
            };
            if (kind == "Excel")
            {
                var excel = ExcelVbeFixture.Start();
                result.fixture = excel; result.Root = excel.Root; result.ProcessId = excel.ProcessId;
                result.Retain = () => excel.PreserveForDiagnosticRecovery = true;
                publish(result);
                result.Data = (command, pairs) =>
                {
                    var request = new Dictionary<string, object> { ["Command"] = command };
                    if (result.Project != null) request["Project"] = result.Project;
                    for (int i = 0; i < pairs.Length; i += 2) request[(string)pairs[i]] = pairs[i + 1];
                    var reply = result.TerminalExcel(excel, request);
                    Assert.IsNotNull(reply, "No terminal bridge response; do not retry.");
                    Assert.AreEqual(true, reply["Ok"], command + ": " + reply["Error"]);
                    return VbeBridgeClient.Object(reply["Data"]);
                };
                var watch = Stopwatch.StartNew(); int observations = 0;
                while (true)
                {
                    result.NativeDispatchUnsettled = true;
                    var identity = excel.ReadOwnedProjectIdentity();
                    result.NativeDispatchUnsettled = false;
                    var rows = (object[])result.TerminalExcel(excel, new Dictionary<string, object> { ["Command"] = "list_projects" })["Data"];
                    string name = Convert.ToString(identity["ProjectName"]);
                    bool ready = Convert.ToInt32(identity["NativeProjectCount"]) == 1 &&
                        Convert.ToInt32(identity["OwnedIdentityMatches"]) == 1 && rows.Length == 1 &&
                        Convert.ToString(VbeBridgeClient.Object(rows[0])["Name"]) == name;
                    bool expired = watch.Elapsed.TotalSeconds >= 30 || observations + 1 >= 128;
                    excel.WriteQualificationEvidence("q028-project-readiness.json", new
                    {
                        State = ready && !expired ? "READY" : (expired ? "FAILED_TIMEOUT" : "OBSERVING"),
                        excel.ProcessId,
                        Identity = identity,
                        BridgeProjectCount = rows.Length,
                        ObservationCount = ++observations,
                        ElapsedMilliseconds = watch.ElapsedMilliseconds,
                        CloseOrAddReplayed = false
                    });
                    if (ready && !expired) { result.Project = name; break; }
                    if (expired)
                        throw new TimeoutException("The exact owned workbook's native project did not become unambiguous; no Close/Add or source mutation will be replayed.");
                    System.Windows.Forms.Application.DoEvents(); Thread.Sleep(100);
                }
                result.Items = command =>
                {
                    var reply = result.TerminalExcel(excel, new Dictionary<string, object> { ["Command"] = command, ["Project"] = result.Project });
                    Assert.IsNotNull(reply); Assert.AreEqual(true, reply["Ok"]);
                    return Array.ConvertAll((object[])reply["Data"], VbeBridgeClient.Object);
                };
            }
            else if (kind == "Outlook")
            {
                var outlook = OutlookVbaTestFixture.Start();
                result.fixture = outlook; result.Root = outlook.Root; result.ProcessId = outlook.ProcessId;
                result.Project = outlook.Project; result.Data = outlook.Data;
                result.Items = outlook.Items;
                // Outlook's existing containment retains unsettled native commands itself.
                result.Retain = () => { };
                publish(result);
            }
            else
            {
                var office = OfficeVbeFixture.Start(kind);
                result.fixture = office; result.Root = office.Root; result.ProcessId = office.ProcessId;
                result.Project = office.Project; result.Data = office.Data;
                result.Items = command => office.Items(command);
                result.Retain = () => office.NativeExecutionUnsettled = true;
                publish(result);
                if (kind == "Access") office.RequireAdapterOnlyCleanup();
            }
            result.NativeDispatchUnsettled = true;
            var projects = result.Items("list_projects");
            result.NativeDispatchUnsettled = false;
            var matched = Array.FindAll(projects, row => Convert.ToString(row["Name"]) == result.Project ||
                string.Equals(VbeProjectHostPath.FromFields(row), result.Project, StringComparison.OrdinalIgnoreCase));
            Assert.AreEqual(1, matched.Length, "The exact owned project must have one scope label.");
            result.projectPath = VbeProjectHostPath.FromFields(matched[0]);
            result.projectName = Convert.ToString(matched[0]["Name"]);
            return result;
        }

        private IDictionary<string, object> TerminalExcel(ExcelVbeFixture excel, IDictionary<string, object> request)
        {
            NativeDispatchUnsettled = true;
            var reply = excel.Command(request);
            if (reply == null || !reply.ContainsKey("Ok") || !(reply["Ok"] is bool))
                throw new InvalidOperationException("No classified terminal bridge response; retain the exact fixture.");
            NativeDispatchUnsettled = false;
            Assert.AreEqual(true, reply["Ok"], Convert.ToString(reply["Error"]));
            return reply;
        }

        internal void TrackOutlookModule(string name)
        {
            (fixture as OutlookVbaTestFixture)?.TrackOwnedModule(name);
        }

        internal void PrepareSyntheticBaseline(string module)
        {
            var office = fixture as OfficeVbeFixture;
            if (office != null && office.Kind == "Publisher")
                office.SaveAdapterBaseline(new[] { module });
        }

        public void Dispose()
        {
            if (retained) throw new InvalidOperationException("Unsettled assistant retains its exact host; no teardown.");
            fixture.Dispose();
        }

        /// <summary>Collects only settled Word testhost temporaries using the existing owned-host diagnostic.</summary>
        internal void CollectSettledWordReferences()
        {
            if (retained || NativeDispatchUnsettled)
                throw new InvalidOperationException("Pending native work must retain its references.");
            (fixture as OfficeVbeFixture)?.CollectSettledWordScopeDiagnostic();
        }
    }
}
