using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Web.Script.Serialization;
using System.Windows.Forms;

namespace VBAi
{
    /// <summary>Disabled-by-default, invocation-bound Word qualification evidence; never establishes host COM readiness.</summary>
    internal sealed class ChatGitModalDiagnostic
    {
        internal const string EnvironmentName = "VBAi_TEST_CHAT_GIT_DIAGNOSTIC_MANIFEST";
        internal const string DirectoryName = "VBAi.ChatGitDiagnostic";
        internal static readonly string[] Phases = { "Started", "ShowModalReturned", "DisposeReturned", "PostHandlerIntent", "PostHandlerCallbackObserved" };

        /// <summary>Observed process/product/windows and the captured invocation scope; current selection is verified separately.</summary>
        internal sealed class Identity
        {
            /// <summary>The canonical document captured for this invocation; not a later selection read.</summary>
            public string DocumentPath { get; set; }
            /// <summary>The actual owning Word process identifier.</summary>
            public int ProcessId { get; set; }
            /// <summary>The original process generation, recorded in round-trip UTC.</summary>
            public string ProcessStartedUtc { get; set; }
            /// <summary>The exact native owning STA thread identifier.</summary>
            public uint ThreadId { get; set; }
            /// <summary>The existing chat HWND, never recreated for observation.</summary>
            public long ChatHandle { get; set; }
            /// <summary>The existing native VBE root HWND.</summary>
            public long RootHandle { get; set; }
            /// <summary>The module version identifier of the assembly actually loaded in the host.</summary>
            public string ProductMvid { get; set; }
            /// <summary>The independently read SHA-256 of that loaded assembly path.</summary>
            public string ProductSha256 { get; set; }
        }

        private readonly string nonce;
        private readonly Identity identity;
        private readonly Func<Identity> observe;
        private readonly Action<string, object> write;
        private readonly List<Exception> errors = new List<Exception>();
        private bool modalReturned, disposeReturned, postAttempted, queueReturned, callbackEntered, failureAttempted;
        /// <summary>Whether the later dispatcher observation was successfully published for this invocation.</summary>
        internal bool Completed { get; private set; }
        /// <summary>Whether any operation, identity or publication error permanently prevents success.</summary>
        internal bool Failed { get { return errors.Count != 0; } }
        /// <summary>The preserved primary or ordered aggregate failure; null when no failure was observed.</summary>
        internal Exception Error { get { return errors.Count == 0 ? null : errors.Count == 1 ? errors[0] : new AggregateException(errors); } }

        /// <summary>Copies the initial identity so later capture-object mutation cannot alter the original evidence.</summary>
        private ChatGitModalDiagnostic(string invocationNonce, Identity actual, Func<Identity> readIdentity, Action<string, object> writer)
        { nonce = invocationNonce; identity = new JavaScriptSerializer().Deserialize<Identity>(new JavaScriptSerializer().Serialize(actual)); observe = readIdentity; write = writer; }

        /// <summary>Reads only an explicitly supplied owned temporary manifest on the live Word chat STA.</summary>
        internal static ChatGitModalDiagnostic BeginFromEnvironment(Control chat, string scope)
        {
            string manifest = Environment.GetEnvironmentVariable(EnvironmentName);
            if (string.IsNullOrWhiteSpace(manifest)) return null;
            string root = RequireManifestPath(manifest, Path.GetTempPath());
            var info = new FileInfo(manifest);
            if (!info.Exists || info.Length < 2 || info.Length > 4096) throw new InvalidOperationException("The chat Git diagnostic manifest is missing or exceeds its bound.");
            using (var process = Process.GetCurrentProcess())
                if (!string.Equals(process.ProcessName, "WINWORD", StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("This diagnostic is restricted to the explicitly owned Word qualification.");
            Func<Identity> capture = () => Capture(chat, scope);
            return Begin(File.ReadAllText(manifest, new UTF8Encoding(false, true)), Path.GetFileName(root), capture,
                (phase, value) => Publish(root, phase, value));
        }

        /// <summary>Validates a versioned per-invocation request before publishing any Started receipt.</summary>
        internal static ChatGitModalDiagnostic Begin(string json, string rootNonce, Func<Identity> readIdentity, Action<string, object> writer)
        {
            if (readIdentity == null || writer == null) throw new ArgumentNullException("Diagnostic dependencies");
            Guid parsed;
            if (!Guid.TryParseExact(rootNonce, "N", out parsed)) throw new ArgumentException("An owned GUID invocation is required.");
            var serializer = new JavaScriptSerializer();
            var request = serializer.DeserializeObject(json) as IDictionary<string, object>;
            if (request == null || request.Count != 3 || !request.ContainsKey("Version") || !(request["Version"] is int) || (int)request["Version"] != 1 ||
                !request.ContainsKey("Nonce") || !string.Equals(request["Nonce"] as string, rootNonce, StringComparison.Ordinal) || !request.ContainsKey("Identity"))
                throw new InvalidOperationException("The diagnostic request version and invocation must be exact.");
            Identity expected = DecodeIdentity(request["Identity"]);
            Identity actual = readIdentity(); RequireSameIdentity(expected, actual);
            var result = new ChatGitModalDiagnostic(rootNonce, actual, readIdentity, writer);
            result.Record("Started", false); // CreateNew publication consumes this invocation; never overwritten.
            return result;
        }

        /// <summary>Preserves show, disposal and diagnostic errors while disposing exactly once.</summary>
        internal void RunModal(Action show, Action dispose)
        {
            if (show == null || dispose == null) throw new ArgumentNullException("Modal dependencies");
            if (modalReturned || disposeReturned || Failed || postAttempted) throw new InvalidOperationException("One modal invocation only.");
            try { show(); modalReturned = true; Record("ShowModalReturned", false); }
            catch (Exception error) { AddError(error); }
            try { dispose(); disposeReturned = true; Record("DisposeReturned", false); }
            catch (Exception error) { AddError(error); }
            if (Failed) throw Error;
        }

        /// <summary>Retains a primary failure even when publication of failure evidence also fails.</summary>
        internal void Fail(Exception error)
        {
            AddError(error);
            if (failureAttempted) return;
            failureAttempted = true;
            try { Record("Failed", false); }
            catch (Exception publication) { AddError(publication); }
        }

        /// <summary>Posts the diagnostic's final action; performs no message pumping or native mutation.</summary>
        internal void SchedulePostHandler(Action<Action> post)
        {
            if (postAttempted) { Fail(new InvalidOperationException("One post-handler publication only.")); return; }
            postAttempted = true;
            if (Failed) return;
            try
            {
                if (!modalReturned || !disposeReturned) throw new InvalidOperationException("Modal and disposal returns are not both proved.");
                Record("PostHandlerIntent", false);
                post(ObservePostHandler);
                queueReturned = true; // Managed flag only; no reentrant operation follows the queue publication.
            }
            catch (Exception error) { Fail(error); }
        }

        /// <summary>Observes a later owning-STA dispatcher turn without claiming COM host readiness.</summary>
        private void ObservePostHandler()
        {
            if (callbackEntered) { Fail(new InvalidOperationException("Deferred callback repeated.")); return; }
            callbackEntered = true;
            if (Failed) return;
            try
            {
                if (!queueReturned || !modalReturned || !disposeReturned) throw new InvalidOperationException("The callback was not deferred beyond the instrumented handler.");
                RequireSameIdentity(identity, observe());
                Record("PostHandlerCallbackObserved", true);
                Completed = true;
            }
            catch (Exception error) { Fail(error); }
        }

        /// <summary>Preserves distinct original exceptions and revokes any success state.</summary>
        private void AddError(Exception error)
        { if (error != null && !errors.Any(existing => ReferenceEquals(existing, error))) errors.Add(error); Completed = false; }

        /// <summary>Publishes the invocation-bound observed stage and all retained failures.</summary>
        private void Record(string phase, bool success)
        { write(phase, new { Version = 1, Nonce = nonce, Phase = phase, Identity = identity, ModalReturned = modalReturned, DisposeReturned = disposeReturned,
            Success = success && !Failed, Errors = errors.Select(error => error.ToString()).ToArray(), Utc = DateTime.UtcNow.ToString("o") }); }

        /// <summary>Decodes exactly the identity fields and rejects implicit string-to-number coercion.</summary>
        internal static Identity DecodeIdentity(object value)
        {
            var fields = value as IDictionary<string, object>;
            string[] names = { "DocumentPath", "ProcessId", "ProcessStartedUtc", "ThreadId", "ChatHandle", "RootHandle", "ProductMvid", "ProductSha256" };
            if (fields == null || fields.Count != names.Length || names.Any(name => !fields.ContainsKey(name))) throw new InvalidOperationException("The identity fields must be exact.");
            foreach (string name in new[] { "ProcessId", "ThreadId", "ChatHandle", "RootHandle" })
                if (!(fields[name] is int) && !(fields[name] is long)) throw new InvalidOperationException("Native identity numbers must be integers.");
            var serializer = new JavaScriptSerializer();
            return serializer.Deserialize<Identity>(serializer.Serialize(fields));
        }

        /// <summary>Rejects a changed process generation, loaded product, canonical scope or native owning windows.</summary>
        internal static void RequireSameIdentity(Identity expected, Identity actual)
        {
            DateTime birth; Guid mvid;
            if (expected == null || actual == null || !Path.IsPathRooted(expected.DocumentPath ?? "") ||
                !string.Equals(expected.DocumentPath, Path.GetFullPath(expected.DocumentPath), StringComparison.OrdinalIgnoreCase) || expected.ProcessId <= 0 || expected.ThreadId == 0 ||
                expected.ChatHandle == 0 || expected.RootHandle == 0 || !DateTime.TryParseExact(expected.ProcessStartedUtc, "o", System.Globalization.CultureInfo.InvariantCulture,
                System.Globalization.DateTimeStyles.RoundtripKind, out birth) || birth.Kind != DateTimeKind.Utc || !Guid.TryParse(expected.ProductMvid, out mvid) ||
                expected.ProductSha256 == null || expected.ProductSha256.Length != 64 || expected.ProductSha256.Any(value => !Uri.IsHexDigit(value)) ||
                !string.Equals(expected.DocumentPath, actual.DocumentPath, StringComparison.OrdinalIgnoreCase) || expected.ProcessId != actual.ProcessId ||
                expected.ProcessStartedUtc != actual.ProcessStartedUtc || expected.ThreadId != actual.ThreadId || expected.ChatHandle != actual.ChatHandle ||
                expected.RootHandle != actual.RootHandle || !string.Equals(expected.ProductMvid, actual.ProductMvid, StringComparison.OrdinalIgnoreCase) ||
                !string.Equals(expected.ProductSha256, actual.ProductSha256, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The exact invocation scope, process generation, owning STA/windows and loaded product are required.");
        }

        /// <summary>Allows only request.json in a fixed, owned GUID child of the caller's local temporary directory.</summary>
        internal static string RequireManifestPath(string path, string temporaryRoot)
        {
            if (string.IsNullOrWhiteSpace(path) || !Path.IsPathRooted(path) || path.StartsWith(@"\\", StringComparison.Ordinal) ||
                !string.Equals(path, Path.GetFullPath(path), StringComparison.Ordinal) || Path.GetFileName(path) != "request.json")
                throw new ArgumentException("A canonical local request.json path is required.");
            string root = Path.GetDirectoryName(path); Guid nonce;
            if (!Guid.TryParseExact(Path.GetFileName(root), "N", out nonce) ||
                !string.Equals(Path.GetDirectoryName(root), Path.Combine(Path.GetFullPath(temporaryRoot), DirectoryName).TrimEnd(Path.DirectorySeparatorChar), StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException("Only an owned diagnostic temporary GUID child is permitted.");
            return root;
        }

        /// <summary>Publishes one closed atomic receipt and then a separate readiness marker, without overwriting.</summary>
        internal static void Publish(string root, string phase, object value)
        {
            if (phase != "Failed" && !Phases.Contains(phase)) throw new ArgumentException("Unknown diagnostic phase.");
            string path = Path.Combine(root, phase + ".json"), temporary = path + ".tmp";
            string json = new JavaScriptSerializer().Serialize(value);
            if (Encoding.UTF8.GetByteCount(json) > 32768) throw new InvalidOperationException("Diagnostic receipt exceeds its bound.");
            using (var file = new FileStream(temporary, FileMode.CreateNew, FileAccess.Write, FileShare.None))
            using (var writer = new StreamWriter(file, new UTF8Encoding(false))) writer.Write(json);
            File.Move(temporary, path);
            using (var ready = new FileStream(path + ".ready", FileMode.CreateNew, FileAccess.Write, FileShare.None)) { }
        }

        /// <summary>Reads the actual process/product and existing owning native windows without invoking COM or creating controls.</summary>
        private static Identity Capture(Control chat, string scope)
        {
            if (chat == null || chat.IsDisposed || !chat.IsHandleCreated || Thread.CurrentThread.GetApartmentState() != ApartmentState.STA)
                throw new InvalidOperationException("The existing live chat on its owning STA is required.");
            uint chatPid, rootPid; IntPtr handle = chat.Handle, root = GetAncestor(handle, 2);
            uint chatTid = GetWindowThreadProcessId(handle, out chatPid), rootTid = GetWindowThreadProcessId(root, out rootPid), caller = GetCurrentThreadId();
            using (var process = Process.GetCurrentProcess())
            {
                if (root == IntPtr.Zero || chatPid != process.Id || rootPid != chatPid || chatTid != caller || rootTid != caller)
                    throw new InvalidOperationException("The exact live owning native STA and root are required.");
                string hash;
                using (var file = new FileStream(typeof(ChatWindow).Assembly.Location, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var sha = SHA256.Create()) hash = BitConverter.ToString(sha.ComputeHash(file)).Replace("-", "");
                return new Identity { DocumentPath = scope, ProcessId = process.Id, ProcessStartedUtc = process.StartTime.ToUniversalTime().ToString("o"), ThreadId = caller,
                    ChatHandle = handle.ToInt64(), RootHandle = root.ToInt64(), ProductMvid = typeof(ChatWindow).Module.ModuleVersionId.ToString("D"), ProductSha256 = hash };
            }
        }
        [DllImport("user32.dll")] private static extern IntPtr GetAncestor(IntPtr handle, uint flag);
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr handle, out uint pid);
        [DllImport("kernel32.dll")] private static extern uint GetCurrentThreadId();
    }
}