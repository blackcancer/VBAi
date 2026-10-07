using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using System.Web.Script.Serialization;

namespace VBAi.Tests.Integration
{
    /// <summary>Validates a complete observed product diagnostic chain before any first host-close attempt.</summary>
    internal static class ChatGitDiagnosticReceipt
    {
        internal static object[] Wait(string root, string nonce, ChatGitModalDiagnostic.Identity expected)
        {
            var watch = Stopwatch.StartNew();
            var serializer = new JavaScriptSerializer();
            return WaitCore(() => watch.ElapsedMilliseconds,
                () => File.Exists(Path.Combine(root, "PostHandlerCallbackObserved.json.ready")),
                () => File.Exists(Path.Combine(root, "Failed.json.ready")),
                () => ChatGitModalDiagnostic.Phases.Select(phase =>
                {
                    string path = Path.Combine(root, phase + ".json");
                    if (!File.Exists(path + ".ready") || new FileInfo(path).Length > 32768) throw new InvalidOperationException("The complete bounded diagnostic chain is required.");
                    return serializer.DeserializeObject(File.ReadAllText(path, new System.Text.UTF8Encoding(false, true)));
                }).ToArray(), () => Thread.Sleep(50), nonce, expected);
        }

        /// <summary>A strict observation deadline includes file reading and validation; I/O errors are never retried.</summary>
        internal static object[] WaitCore(Func<long> elapsed, Func<bool> ready, Func<bool> failed, Func<object[]> read,
            Action pause, string nonce, ChatGitModalDiagnostic.Identity expected)
        {
            if (elapsed == null || ready == null || failed == null || read == null || pause == null) throw new ArgumentNullException("Receipt dependencies");
            while (true)
            {
                if (elapsed() >= 15000) throw new TimeoutException("The installed post-handler observation exceeded its deadline; no Close is authorized.");
                if (failed()) throw new InvalidOperationException("A failed diagnostic cannot authorize Close.");
                if (ready()) break;
                pause();
            }
            if (elapsed() >= 15000) throw new TimeoutException("A terminal appearing after the deadline cannot authorize Close.");
            object[] chain = read();
            Validate(chain, nonce, expected);
            if (failed()) throw new InvalidOperationException("Diagnostic failure appeared before the first Close authorization.");
            if (elapsed() >= 15000) throw new TimeoutException("Diagnostic reading and validation exceeded the observation deadline.");
            return chain;
        }
        internal static void Validate(object[] chain, string nonce, ChatGitModalDiagnostic.Identity expected)
        {
            Guid guid;
            if (chain == null || chain.Length != ChatGitModalDiagnostic.Phases.Length || !Guid.TryParseExact(nonce, "N", out guid))
                throw new InvalidOperationException("One complete invocation chain is required.");
            DateTime previous = DateTime.MinValue;
            var serializer = new JavaScriptSerializer();
            for (int index = 0; index < chain.Length; index++)
            {
                var row = chain[index] as IDictionary<string, object>;
                if (row == null || row.Count != 9 || !row.ContainsKey("Version") || !(row["Version"] is int) || (int)row["Version"] != 1 ||
                    !row.ContainsKey("Nonce") || !string.Equals(row["Nonce"] as string, nonce, StringComparison.Ordinal) ||
                    !row.ContainsKey("Phase") || !string.Equals(row["Phase"] as string, ChatGitModalDiagnostic.Phases[index], StringComparison.Ordinal) ||
                    !row.ContainsKey("Errors") || !(row["Errors"] is object[]) || ((object[])row["Errors"]).Length != 0 ||
                    !row.ContainsKey("Identity") || !(row["Identity"] is IDictionary<string, object>) ||
                    !row.ContainsKey("ModalReturned") || !(row["ModalReturned"] is bool) || (bool)row["ModalReturned"] != (index >= 1) ||
                    !row.ContainsKey("DisposeReturned") || !(row["DisposeReturned"] is bool) || (bool)row["DisposeReturned"] != (index >= 2) ||
                    !row.ContainsKey("Success") || !(row["Success"] is bool) || (bool)row["Success"] != (index == chain.Length - 1) || !row.ContainsKey("Utc"))
                    throw new InvalidOperationException("The diagnostic chain is incomplete, foreign or failed.");
                ChatGitModalDiagnostic.RequireSameIdentity(expected, ChatGitModalDiagnostic.DecodeIdentity(row["Identity"]));
                DateTime utc;
                if (!DateTime.TryParseExact(row["Utc"] as string, "o", System.Globalization.CultureInfo.InvariantCulture,
                    System.Globalization.DateTimeStyles.RoundtripKind, out utc) || utc.Kind != DateTimeKind.Utc || utc < previous)
                    throw new InvalidOperationException("Diagnostic phase timestamps must be ordered UTC observations.");
                previous = utc;
            }
        }
    }
}