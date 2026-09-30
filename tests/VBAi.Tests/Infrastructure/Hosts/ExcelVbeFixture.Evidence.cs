using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.ExceptionServices;
using System.Text;
using System.Web.Script.Serialization;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Integration
{
    internal sealed partial class ExcelVbeFixture
    {
        /// <summary>Opt-in evidence retention; the ordinary temporary fixture stays unchanged.</summary>
        private bool retainEvidence;
        /// <summary>Maximum request/response records per synthetic fixture.</summary>
        private const int MaximumCommandRecords = 128;
        /// <summary>Observed protocol calls, including records beyond the retention bound.</summary>
        private int commandSequence;
        /// <summary>Latest startup observation, preserved through a later startup failure.</summary>
        private readonly Dictionary<string, object> startupEvidence = new Dictionary<string, object>();

        /// <summary>Records actual owned startup identity and native bridge observation without another command.</summary>
        private void RecordStartup(string phase, int[] existingIds, IDictionary<string, object> status = null, Exception error = null)
        {
            if (!retainEvidence) return;
            if (!startupEvidence.ContainsKey("StartedUtc"))
            {
                startupEvidence["StartedUtc"] = DateTime.UtcNow.ToString("o");
                var test = new StackTrace().GetFrames()?.Select(frame => frame.GetMethod())
                    .FirstOrDefault(method => method != null && method.GetCustomAttributes(typeof(TestMethodAttribute), true).Length != 0);
                startupEvidence["Scenario"] = test == null ? null : test.DeclaringType.FullName + "." + test.Name;
            }
            startupEvidence["ObservedUtc"] = DateTime.UtcNow.ToString("o");
            startupEvidence["Phase"] = phase;
            startupEvidence["FixtureRoot"] = Root;
            startupEvidence["ProcessId"] = ProcessId;
            startupEvidence["Owned"] = owned;
            startupEvidence["ExistingExcelProcessIds"] = existingIds;
            startupEvidence["ExpectedAssemblyMvid"] = typeof(VbeSession).Module.ModuleVersionId.ToString("D");
            if (ownedProcess != null && !startupEvidence.ContainsKey("HostExecutable"))
            {
                string executable = ownedImagePath != null ? ownedImagePath() : ownedProcess.MainModule.FileName;
                var version = FileVersionInfo.GetVersionInfo(executable);
                startupEvidence["HostExecutable"] = executable;
                startupEvidence["HostFileVersion"] = version.FileVersion;
                startupEvidence["HostProductVersion"] = version.ProductVersion;
                startupEvidence["HostStartedUtc"] = ownedProcess.StartTime.ToUniversalTime().ToString("o");
                startupEvidence["HostVersion"] = Convert.ToString(((dynamic)application).Version);
                startupEvidence["HostBuild"] = Convert.ToString(((dynamic)application).Build);
            }
            if (workbook != null && !startupEvidence.ContainsKey("InitialWorkbook"))
                startupEvidence["InitialWorkbook"] = new {
                    Name = Convert.ToString(((dynamic)workbook).Name),
                    FullName = Convert.ToString(((dynamic)workbook).FullName),
                    Saved = Convert.ToBoolean(((dynamic)workbook).Saved)
                };
            if (status != null)
            {
                startupEvidence["BridgeStatus"] = status;
                if (status.TryGetValue("Data", out var data) && data is IDictionary<string, object> state &&
                    state.TryGetValue("AssemblyModuleVersionId", out var mvid)) startupEvidence["LoadedAssemblyMvid"] = mvid;
            }
            if (error != null) startupEvidence["Error"] = error.ToString();
            WriteEvidence("startup.json", startupEvidence);
        }

        /// <summary>Records each synthetic request once, with no replay after an uncertain response or evidence error.</summary>
        private IDictionary<string, object> RecordCommand(object request, Func<IDictionary<string, object>> execute)
        {
            if (!retainEvidence) return execute();
            int sequence = ++commandSequence;
            if (sequence > MaximumCommandRecords) return execute();
            var record = new Dictionary<string, object> {
                ["Sequence"] = sequence, ["ProcessId"] = ProcessId,
                ["StartedUtc"] = DateTime.UtcNow.ToString("o"), ["Request"] = BoundedJson(request, 16384)
            };
            IDictionary<string, object> response = null;
            Exception failure = null;
            try { response = execute(); }
            catch (Exception error) { failure = error; }
            record["FinishedUtc"] = DateTime.UtcNow.ToString("o");
            record["Error"] = failure?.ToString();
            try
            {
                record["Response"] = BoundedJson(response, 65536);
                WriteEvidence("command-" + sequence.ToString("D4") + ".json", record);
            }
            catch (Exception evidence)
            {
                if (failure != null) throw new AggregateException("The Excel command and its evidence write failed; the request was not replayed.", failure, evidence);
                throw;
            }
            if (failure != null) ExceptionDispatchInfo.Capture(failure).Throw();
            return response;
        }

        /// <summary>Bounds retained synthetic payload text while explicitly recording truncation.</summary>
        private static object BoundedJson(object value, int limit)
        {
            string text = new JavaScriptSerializer { MaxJsonLength = 10 * 1024 * 1024 }.Serialize(value);
            return new { Json = text.Length <= limit ? text : text.Substring(0, limit), OriginalCharacters = text.Length, Truncated = text.Length > limit };
        }

        /// <summary>Writes durable UTF-8 evidence; an opted-in write failure remains a visible qualification failure.</summary>
        private void WriteEvidence(string name, object value)
        {
            string text = new JavaScriptSerializer { MaxJsonLength = 10 * 1024 * 1024 }.Serialize(value);
            System.IO.File.WriteAllText(Path.Combine(Root, name), text, new UTF8Encoding(false));
        }
    }
}
