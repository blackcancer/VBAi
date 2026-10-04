using System;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Integration
{
    /// <summary>Adapts existing owned-host fixtures without bypassing their containment or shutdown guards.</summary>
    internal sealed class OllamaOfficeHost : IDisposable
    {
        internal delegate IDictionary<string, object> ReadData(string command, params object[] pairs);
        internal ReadData Data;
        internal Func<string, IDictionary<string, object>[]> Items;
        internal string Project, Root, Label;
        internal int ProcessId;
        internal Action Retain;
        private IDisposable fixture;
        private bool retained;
        private static readonly List<OllamaOfficeHost> Retained = new List<OllamaOfficeHost>();

        internal static OllamaOfficeHost Start(string kind)
        {
            var result = new OllamaOfficeHost();
            if (kind == "Excel")
            {
                var excel = ExcelVbeFixture.Start();
                result.fixture = excel; result.Root = excel.Root; result.ProcessId = excel.ProcessId;
                result.Retain = () => excel.PreserveForDiagnosticRecovery = true;
                result.Data = (command, pairs) => {
                    var request = new Dictionary<string, object> { ["Command"] = command };
                    if (result.Project != null) request["Project"] = result.Project;
                    for (int i = 0; i < pairs.Length; i += 2) request[(string)pairs[i]] = pairs[i + 1];
                    var reply = excel.Command(request);
                    Assert.IsNotNull(reply, "No terminal bridge response; do not retry.");
                    Assert.AreEqual(true, reply["Ok"], command + ": " + reply["Error"]);
                    return VbeBridgeClient.Object(reply["Data"]);
                };
                var rows = (object[])excel.Command("list_projects")["Data"];
                Assert.AreEqual(1, rows.Length, "The disposable Excel instance must expose exactly one project.");
                result.Project = Convert.ToString(VbeBridgeClient.Object(rows[0])["Name"]);
                result.Items = command => {
                    var reply = excel.Command(new Dictionary<string, object> { ["Command"] = command, ["Project"] = result.Project });
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
            }
            else
            {
                var office = OfficeVbeFixture.Start(kind);
                result.fixture = office; result.Root = office.Root; result.ProcessId = office.ProcessId;
                result.Project = office.Project; result.Data = office.Data;
                result.Items = command => office.Items(command);
                result.Retain = () => office.NativeExecutionUnsettled = true;
            }
            var retainFixture = result.Retain;
            result.Retain = () => { result.retained = true; Retained.Add(result); retainFixture(); };
            var projects = result.Items("list_projects");
            var matched = Array.FindAll(projects, row => Convert.ToString(row["Name"]) == result.Project ||
                string.Equals(VbeProjectHostPath.FromFields(row), result.Project, StringComparison.OrdinalIgnoreCase));
            Assert.AreEqual(1, matched.Length, "The exact owned project must have one scope label.");
            string path = VbeProjectHostPath.FromFields(matched[0]);
            bool saved = !string.IsNullOrWhiteSpace(path) && System.IO.Path.IsPathRooted(path);
            result.Label = Convert.ToString(matched[0]["Name"]) + " · " + (saved ? System.IO.Path.GetFileName(path) : UiText.Get("unsaved document"));
            return result;
        }

        internal void TrackOutlookModule(string name)
        {
            (fixture as OutlookVbaTestFixture)?.TrackOwnedModule(name);
        }

        public void Dispose()
        {
            if (retained) throw new InvalidOperationException("Unsettled assistant retains its exact host; no teardown.");
            fixture.Dispose();
        }
    }
}
