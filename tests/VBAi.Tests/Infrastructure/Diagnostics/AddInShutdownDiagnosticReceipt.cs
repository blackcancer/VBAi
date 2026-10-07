using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Web.Script.Serialization;

namespace VBAi.Tests.Integration
{
    /// <summary>Pure bounded validation of installed cleanup observations, separate from original-process exit evidence.</summary>
    internal static class AddInShutdownDiagnosticReceipt
    {
        internal sealed class Invocation
        {
            public string Id, Parent, EntryPoint, Outcome;
            public readonly IDictionary<string, string> Stages = new Dictionary<string, string>();
            public readonly IList<string> ClearedReferences = new List<string>();
            internal string ActiveStage;
            internal int LastOperationRank = -1;
            internal bool RequiresFaultTerminal;
        }

        private static readonly string[] Cleared = { "editorNavigation", "editorWorkspace", "modernEditor", "testExplorerWindow",
            "nativeTestWindow", "nativeTestControl", "testExplorerService", "crashReporter", "menu", "chat", "nativeChatWindow",
            "nativeChatControl", "addIn", "server", "dispatcher", "vbe" };
        private static readonly string[] DisposeStages = { "VbeNativeTheme.Disconnect", "WriteLog.ThemeDeferred", "editorNavigation.Dispose", "editorWorkspace.Dispose",
            "modernEditor.Dispose", "nativeTestControl.Detach", "testExplorerWindow.Dispose", "nativeTestWindow.Close", "testExplorerService.Dispose", "StopUpdateCheck",
            "crashReporter.Dispose", "menu.Dispose", "nativeChatControl.Detach", "chat.Dispose", "nativeChatWindow.Close", "server.Dispose", "dispatcher.Dispose" };
        private static readonly string[] OperationOrder = { "VbeNativeTheme.Disconnect", "WriteLog.ThemeDeferred", "editorNavigation.Dispose", "clear:editorNavigation",
            "editorWorkspace.Dispose", "clear:editorWorkspace", "modernEditor.Dispose", "clear:modernEditor", "nativeTestControl.Detach", "testExplorerWindow.Dispose",
            "clear:testExplorerWindow", "nativeTestWindow.Close", "clear:nativeTestWindow", "clear:nativeTestControl", "testExplorerService.Dispose", "clear:testExplorerService",
            "StopUpdateCheck", "crashReporter.Dispose", "clear:crashReporter", "menu.Dispose", "clear:menu", "nativeChatControl.Detach", "chat.Dispose", "clear:chat",
            "nativeChatWindow.Close", "clear:nativeChatWindow", "clear:nativeChatControl", "clear:addIn", "server.Dispose", "clear:server", "dispatcher.Dispose", "clear:dispatcher", "clear:vbe" };

        /// <summary>Validates every sequence/identity/parent/operation terminal, without treating returned cleanup as host exit.</summary>
        internal static Invocation[] Validate(object[] records, string nonce, AddInShutdownDiagnostic.Identity expected)
        {
            Guid guid;
            if (!Guid.TryParseExact(nonce, "N", out guid) || records == null || records.Length == 0 || records.Length > AddInShutdownDiagnostic.MaximumEvents)
                throw new InvalidOperationException("A bounded complete shutdown chain is required.");
            AddInShutdownDiagnostic.RequireSameIdentity(expected, expected);
            var invocations = new Dictionary<string, Invocation>();
            var stack = new List<Invocation>();
            long elapsed = -1;
            string[] schema = { "Version", "Nonce", "Identity", "InvocationId", "ParentInvocationId", "EntryPoint", "Sequence", "Phase", "Stage", "Thread", "Utc", "ElapsedTicks", "Error", "FaultCaught", "ErrorIsComException" };
            for (int index = 0; index < records.Length; index++)
            {
                var row = records[index] as IDictionary<string, object>;
                if (row == null || row.Count != schema.Length || schema.Any(name => !row.ContainsKey(name)) ||
                    !(row["Version"] is int) || (int)row["Version"] != 1 || row["Nonce"] as string != nonce ||
                    !(row["Sequence"] is int) || (int)row["Sequence"] != index + 1)
                    throw new InvalidOperationException("Shutdown sequence/schema/nonce is incomplete or foreign.");
                AddInShutdownDiagnostic.RequireSameIdentity(expected, AddInShutdownDiagnostic.DecodeIdentity(row["Identity"]));
                var thread = row["Thread"] as IDictionary<string, object>;
                if (thread == null || thread.Count != 3 || !thread.ContainsKey("ManagedThreadId") || !thread.ContainsKey("NativeThreadId") || !thread.ContainsKey("Apartment") ||
                    !(thread["ManagedThreadId"] is int) || (int)thread["ManagedThreadId"] <= 0 || Integer(thread["NativeThreadId"]) != expected.ThreadId || thread["Apartment"] as string != "STA")
                    throw new InvalidOperationException("Shutdown observation was not made on the expected native STA.");
                DateTime utc;
                if (!DateTime.TryParseExact(row["Utc"] as string, "o", CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out utc) || utc.Kind != DateTimeKind.Utc ||
                    Integer(row["ElapsedTicks"]) < 0 || Integer(row["ElapsedTicks"]) < elapsed) throw new InvalidOperationException("Shutdown monotonic time is invalid.");
                elapsed = Integer(row["ElapsedTicks"]);
                string id = row["InvocationId"] as string, parent = row["ParentInvocationId"] as string, entry = row["EntryPoint"] as string;
                string phase = row["Phase"] as string, stage = row["Stage"] as string, error = row["Error"] as string;
                if (!Guid.TryParseExact(id, "N", out guid) || (row["ParentInvocationId"] != null && (parent == null || !Guid.TryParseExact(parent, "N", out guid))) ||
                    (row["Stage"] != null && string.IsNullOrWhiteSpace(stage)) ||
                    (phase == "Fault" ? string.IsNullOrWhiteSpace(error) || !(row["FaultCaught"] is bool) || !(row["ErrorIsComException"] is bool) :
                        row["Error"] != null || row["FaultCaught"] != null || row["ErrorIsComException"] != null))
                    throw new InvalidOperationException("Shutdown invocation/event types are invalid.");
                Invocation invocation;
                if (stage == null && phase == "Entry")
                {
                    if (invocations.ContainsKey(id) || invocations.Count >= AddInShutdownDiagnostic.MaximumInvocations ||
                        (entry != "OnBeginShutdown" && entry != "OnDisconnection" && entry != "Dispose") ||
                        (parent == null ? stack.Count != 0 : stack.Count == 0 || stack.Last().Id != parent || stack.Last().RequiresFaultTerminal))
                        throw new InvalidOperationException("Shutdown invocation entry/parent is invalid.");
                    invocation = new Invocation { Id = id, Parent = parent, EntryPoint = entry };
                    invocations.Add(id, invocation); stack.Add(invocation); continue;
                }
                if (!invocations.TryGetValue(id, out invocation) || stack.Count == 0 || stack.Last() != invocation || invocation.Parent != parent || invocation.EntryPoint != entry || invocation.Outcome != null)
                    throw new InvalidOperationException("Shutdown event has no active exact invocation.");
                if (invocation.RequiresFaultTerminal && (stage != null || phase != "Fault")) throw new InvalidOperationException("Propagated cleanup fault cannot resume the original cleanup.");
                if (stage == null)
                {
                    if ((phase != "Returned" && phase != "Fault") || invocation.ActiveStage != null || (phase == "Fault" && (bool)row["FaultCaught"])) throw new InvalidOperationException("Invocation terminal leaves an unresolved operation.");
                    if (phase == "Returned" && entry == "Dispose" &&
                        (!invocation.ClearedReferences.SequenceEqual(Cleared) || !invocation.Stages.ContainsKey("VbeNativeTheme.Disconnect") ||
                         !invocation.Stages.TryGetValue("StopUpdateCheck", out var update) || update != "Returned"))
                        throw new InvalidOperationException("Returned Dispose lacks its required observed cleanup/assignments.");
                    if (phase == "Returned" && entry != "Dispose" &&
                        invocations.Values.Count(child => child.Parent == id && child.EntryPoint == "Dispose" && child.Outcome == "Returned") != 1)
                        throw new InvalidOperationException("Returned callback lacks its one completed Dispose child.");
                    invocation.Outcome = phase; stack.RemoveAt(stack.Count - 1); continue;
                }
                if (phase == "ManagedReferenceCleared")
                {
                    if (entry != "Dispose" || invocation.ActiveStage != null || !Cleared.Contains(stage) || invocation.ClearedReferences.Contains(stage) ||
                        invocation.ClearedReferences.Count >= Cleared.Length || Cleared[invocation.ClearedReferences.Count] != stage)
                        throw new InvalidOperationException("A reference assignment is duplicate, out of order or mislabeled.");
                    RequireOperationOrder(invocation, "clear:" + stage);
                    invocation.ClearedReferences.Add(stage); continue;
                }
                bool knownStage = entry == "Dispose" ? DisposeStages.Contains(stage) : stage == "CleanupTemporaryToolbarCommands";
                if (!knownStage) throw new InvalidOperationException("Unknown existing cleanup operation.");
                if (phase == "Entry")
                {
                    if (invocation.ActiveStage != null || invocation.Stages.ContainsKey(stage)) throw new InvalidOperationException("Cleanup operation is repeated or unresolved.");
                    if (entry == "Dispose") RequireOperationOrder(invocation, stage);
                    invocation.ActiveStage = stage; invocation.Stages.Add(stage, "Entry");
                }
                else
                {
                    if ((phase != "Returned" && phase != "Fault") || invocation.ActiveStage != stage) throw new InvalidOperationException("Cleanup terminal lacks exact entry.");
                    if (phase == "Fault")
                    {
                        bool caught = (bool)row["FaultCaught"];
                        if (caught && stage != "CleanupTemporaryToolbarCommands" && stage != "VbeNativeTheme.Disconnect" && stage != "WriteLog.ThemeDeferred" && stage != "nativeChatWindow.Close" &&
                            (stage != "nativeTestWindow.Close" || !(bool)row["ErrorIsComException"])) throw new InvalidOperationException("The actual cleanup catch policy cannot absorb this fault.");
                        invocation.RequiresFaultTerminal = !caught;
                    }
                    invocation.Stages[stage] = phase; invocation.ActiveStage = null;
                }
            }
            if (stack.Count != 0 || invocations.Values.Any(invocation => invocation.Outcome == null)) throw new InvalidOperationException("Shutdown trace lacks a terminal.");
            return invocations.Values.ToArray();
        }

        private static void RequireOperationOrder(Invocation invocation, string operation)
        {
            int rank = Array.IndexOf(OperationOrder, operation);
            if (rank < 0 || rank <= invocation.LastOperationRank) throw new InvalidOperationException("Cleanup operations and reference assignments changed their original order.");
            invocation.LastOperationRank = rank;
        }

        private static long Integer(object value)
        { if (value is int) return (int)value; if (value is long) return (long)value; throw new InvalidOperationException("An exact diagnostic integer is required."); }

        /// <summary>Reads once after the unchanged native exit observation; no polling or later rehabilitation.</summary>
        internal static object Read(string output, string nonce, AddInShutdownDiagnostic.Identity expected)
        {
            AddInShutdownDiagnostic.RequireNoReparse(output);
            var paths = Directory.EnumerateFiles(output, "*", SearchOption.TopDirectoryOnly).Take(AddInShutdownDiagnostic.MaximumEvents * 2 + 1).ToArray();
            if (paths.Length == 0 || paths.Length > AddInShutdownDiagnostic.MaximumEvents * 2 || Directory.EnumerateDirectories(output).Take(1).Any())
                throw new InvalidOperationException("Shutdown output is absent or exceeds its bound.");
            var json = new JavaScriptSerializer();
            var records = new List<object>();
            var proofs = new List<object>();
            int terminals = 0;
            foreach (string path in paths.Where(file => Path.GetExtension(file) == ".json").OrderBy(file => file, StringComparer.Ordinal))
            {
                AddInShutdownDiagnostic.RequireNoReparse(path);
                string name = Path.GetFileName(path);
                if (name != (records.Count + 1).ToString("D6", CultureInfo.InvariantCulture) + ".json" || new FileInfo(path).Length > AddInShutdownDiagnostic.MaximumBytes)
                    throw new InvalidOperationException("Shutdown receipt file is missing, foreign or oversized.");
                string text = File.ReadAllText(path, new UTF8Encoding(false, true));
                var row = json.DeserializeObject(text) as IDictionary<string, object>;
                if (row == null || text != json.Serialize(row)) throw new InvalidOperationException("Canonical shutdown receipts are required.");
                records.Add(row);
                bool terminal = row.TryGetValue("Stage", out var stage) && stage == null && row.TryGetValue("Phase", out var phase) && (phase as string == "Returned" || phase as string == "Fault");
                string ready = path + ".ready";
                if (terminal)
                {
                    AddInShutdownDiagnostic.RequireNoReparse(ready);
                    if (!File.Exists(ready) || new FileInfo(ready).Length != 0) throw new InvalidOperationException("Shutdown terminal publication was not observed complete.");
                    terminals++;
                }
                else if (File.Exists(ready)) throw new InvalidOperationException("A nonterminal shutdown marker is mislabeled ready.");
                using (var file = File.OpenRead(path))
                using (var hash = SHA256.Create()) proofs.Add(new { Path = path, Sha256 = BitConverter.ToString(hash.ComputeHash(file)).Replace("-", "") });
            }
            if (paths.Length != records.Count + terminals) throw new InvalidOperationException("Partial, extra or temporary shutdown files remain.");
            Invocation[] invocations = Validate(records.ToArray(), nonce, expected);
            return new
            {
                State = "VALID_CLEANUP_OBSERVATIONS_ONLY",
                Identity = expected,
                Invocations = invocations,
                EvidenceFiles = proofs,
                NativeRcwReleaseProven = false,
                HostExitProven = false
            };
        }
    }
}
