using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Web.Script.Serialization;

namespace VBAi
{
    /// <summary>Opt-in observations of existing cleanup calls; neither owns COM resources nor proves process exit.</summary>
    internal sealed class AddInShutdownDiagnostic
    {
        internal const string EnvironmentName = "VBAi_TEST_ADDIN_SHUTDOWN_ROOT";
        internal const string DirectoryName = "VBAi.AddInShutdownDiagnostic";
        internal const int MaximumInvocations = 16, MaximumEvents = 2048, MaximumBytes = 16384;

        /// <summary>The actual process generation, loaded file and expected native callback thread; contains no project lease.</summary>
        internal sealed class Identity
        {
            public int ProcessId { get; set; }
            public string ProcessStartedUtc { get; set; }
            public string HostImagePath { get; set; }
            public string ProductPath { get; set; }
            public string ProductMvid { get; set; }
            public string ProductSha256 { get; set; }
            public uint ThreadId { get; set; }
        }

        /// <summary>Caller values actually observed at an individual marker, independently of a requested identity.</summary>
        internal sealed class ThreadIdentity
        {
            public int ManagedThreadId { get; set; }
            public uint NativeThreadId { get; set; }
            public string Apartment { get; set; }
        }

        /// <summary>Bounded per-request publication state; no delegates capture an add-in, window or COM object.</summary>
        private sealed class Session
        {
            internal string Nonce;
            internal Identity Identity;
            internal Func<ThreadIdentity> ObserveThread;
            internal Action<int, object, bool> Publish;
            internal readonly Stopwatch Clock = Stopwatch.StartNew();
            internal int Invocations, Sequence;
            internal Exception Error;
            internal readonly List<AddInShutdownDiagnostic> Active = new List<AddInShutdownDiagnostic>();
        }

        private static readonly object RuntimeLock = new object();
        private static Session runtimeSession;
        private static string runtimeRequest;
        private readonly Session session;
        private readonly string entryPoint, parent;
        private string activeStage;
        private bool terminal;
        internal string InvocationId { get; private set; }
        internal bool DiagnosticFailed { get { return session.Error != null; } }
        internal Exception PublicationError { get { return session.Error; } }

        private AddInShutdownDiagnostic(Session shared, string entry, string parentId)
        { session = shared; entryPoint = entry; parent = parentId; InvocationId = Guid.NewGuid().ToString("N"); }

        /// <summary>Returns immediately when disabled; an admission or publication error never prevents original cleanup.</summary>
        internal static AddInShutdownDiagnostic TryBeginFromEnvironment(string entryPoint)
        {
            string requestedRoot = Environment.GetEnvironmentVariable(EnvironmentName);
            if (string.IsNullOrWhiteSpace(requestedRoot)) return null;
            try
            {
                string root = RequireRoot(requestedRoot, Path.GetTempPath());
                Identity actual = CaptureIdentity();
                string request = Path.Combine(root, "request-" + actual.ProcessId.ToString(CultureInfo.InvariantCulture) + ".json");
                RequireNoReparse(request);
                var file = new FileInfo(request);
                if (!file.Exists) return null; // Startup can fail before the owned fixture has published Ready.
                if (file.Length < 2 || file.Length > MaximumBytes) throw new InvalidOperationException("Shutdown request exceeds its bound.");
                string json = File.ReadAllText(request, new UTF8Encoding(false, true));
                string nonce = Path.GetFileName(root);
                Identity expected = DecodeRequest(json, nonce);
                RequireSameIdentity(expected, actual);
                lock (RuntimeLock)
                {
                    if (runtimeSession == null)
                    {
                        string output = Path.Combine(root, "process-" + actual.ProcessId.ToString(CultureInfo.InvariantCulture));
                        RequireNoReparse(output);
                        if (Directory.Exists(output)) throw new InvalidOperationException("Shutdown output already exists; never overwrite another process generation.");
                        Directory.CreateDirectory(output);
                        runtimeRequest = request;
                        runtimeSession = CreateSession(nonce, actual, CaptureThread,
                            (sequence, value, isTerminal) => Publish(output, sequence, value, isTerminal));
                    }
                    else
                    {
                        if (!string.Equals(request, runtimeRequest, StringComparison.OrdinalIgnoreCase))
                            throw new InvalidOperationException("A shutdown session cannot change its request path.");
                        RequireSameIdentity(runtimeSession.Identity, expected);
                    }
                    string parent = runtimeSession.Active.Count == 0 ? null : runtimeSession.Active.Last().InvocationId;
                    return NewInvocation(runtimeSession, entryPoint, parent);
                }
            }
            catch (Exception error)
            {
                // Logging is also observational: a substituted logger must not replace a cleanup error.
                try { AddIn.WriteLog("Shutdown diagnostic admission failed: " + error.GetType().FullName); } catch { }
                return null;
            }
        }

        /// <summary>Pure admission seam used by mirror tests and the independently bound native receipt reader.</summary>
        internal static AddInShutdownDiagnostic Begin(string json, string nonce, string entryPoint,
            Func<Identity> observe, Func<ThreadIdentity> observeThread, Action<int, object, bool> publish)
        {
            if (observe == null || observeThread == null || publish == null) throw new ArgumentNullException("Shutdown diagnostic dependencies");
            Identity expected = DecodeRequest(json, nonce), actual = observe();
            RequireSameIdentity(expected, actual);
            return NewInvocation(CreateSession(nonce, actual, observeThread, publish), entryPoint, null);
        }

        private static Session CreateSession(string nonce, Identity actual, Func<ThreadIdentity> observe,
            Action<int, object, bool> publish)
        {
            var json = new JavaScriptSerializer();
            return new Session { Nonce = nonce, Identity = json.Deserialize<Identity>(json.Serialize(actual)),
                ObserveThread = observe, Publish = publish };
        }

        private static AddInShutdownDiagnostic NewInvocation(Session shared, string entry, string parent)
        {
            try
            {
                if (shared.Error != null) return null;
                if (entry != "OnBeginShutdown" && entry != "OnDisconnection" && entry != "Dispose")
                    throw new ArgumentException("Unknown shutdown entry point.");
                if (parent == null ? shared.Active.Count != 0 : shared.Active.Count == 0 || shared.Active.Last().InvocationId != parent)
                    throw new InvalidOperationException("Shutdown child requires the currently active exact parent.");
                if (++shared.Invocations > MaximumInvocations) throw new InvalidOperationException("Shutdown invocation bound exceeded.");
                var result = new AddInShutdownDiagnostic(shared, entry, parent);
                result.Record("Entry", null, null);
                shared.Active.Add(result);
                return result;
            }
            catch (Exception error) { shared.Error = shared.Error ?? error; return null; }
        }

        /// <summary>Creates a distinct linked invocation without retaining the resource being disposed.</summary>
        internal AddInShutdownDiagnostic Child(string childEntryPoint)
        { return NewInvocation(session, childEntryPoint, InvocationId); }

        /// <summary>Marks an operation about to run; no native or managed resource is captured.</summary>
        internal void Enter(string stage)
        {
            if (DiagnosticFailed) return;
            if (terminal || activeStage != null || string.IsNullOrWhiteSpace(stage)) { Fail(new InvalidOperationException("Invalid shutdown stage entry.")); return; }
            activeStage = stage; Record("Entry", stage, null);
        }

        /// <summary>Marks only the normal return of the previously entered existing operation.</summary>
        internal void Returned(string stage)
        {
            if (DiagnosticFailed) return;
            if (terminal || activeStage != stage) { Fail(new InvalidOperationException("Shutdown return has no exact entered stage.")); return; }
            Record("Returned", stage, null); activeStage = null;
        }

        /// <summary>Records an error absorbed by an existing catch without turning it into a normal operation return.</summary>
        internal void CaughtFault(string stage, Exception error)
        {
            if (DiagnosticFailed) return;
            if (terminal || activeStage != stage) { Fail(new InvalidOperationException("Shutdown caught fault has no exact stage.")); return; }
            Record("Fault", stage, error, true); activeStage = null;
        }

        /// <summary>Observes an existing null assignment, without claiming RCW release or native Close.</summary>
        internal void ManagedReferenceCleared(string field)
        {
            if (DiagnosticFailed) return;
            if (terminal || activeStage != null || string.IsNullOrWhiteSpace(field)) { Fail(new InvalidOperationException("Invalid cleared-reference observation.")); return; }
            Record("ManagedReferenceCleared", field, null);
        }

        /// <summary>Records propagated failure; the caller rethrows the same original exception.</summary>
        internal void Fault(Exception error)
        {
            if (DiagnosticFailed) return;
            if (terminal || error == null) { Fail(new InvalidOperationException("Invalid shutdown invocation fault.")); return; }
            if (activeStage != null) { Record("Fault", activeStage, error); activeStage = null; }
            Finish("Fault", error);
        }

        /// <summary>Only states that the existing invocation returned; host exit and native release remain independent.</summary>
        internal void Complete()
        {
            if (DiagnosticFailed) return;
            if (terminal || activeStage != null) { Fail(new InvalidOperationException("Shutdown completion has an unresolved stage.")); return; }
            Finish("Returned", null);
        }

        private void Finish(string phase, Exception error)
        {
            if (session.Active.Count == 0 || session.Active.Last() != this) { Fail(new InvalidOperationException("Shutdown invocation cannot finish ahead of its active child.")); return; }
            Record(phase, null, error); terminal = true;
            session.Active.RemoveAt(session.Active.Count - 1);
        }

        private void Fail(Exception error) { session.Error = session.Error ?? error; }

        /// <summary>Stops all subsequent publication on the first diagnostic error; never changes cleanup propagation.</summary>
        private void Record(string phase, string stage, Exception error, bool caught = false)
        {
            if (DiagnosticFailed) return;
            try
            {
                ThreadIdentity thread = session.ObserveThread();
                if (thread == null || thread.ManagedThreadId <= 0 || thread.NativeThreadId != session.Identity.ThreadId || thread.Apartment != "STA")
                    throw new InvalidOperationException("Shutdown observation left its owning native STA.");
                int sequence = ++session.Sequence;
                if (sequence > MaximumEvents) throw new InvalidOperationException("Shutdown event bound exceeded.");
                session.Publish(sequence, new { Version = 1, session.Nonce, session.Identity, InvocationId, ParentInvocationId = parent,
                    EntryPoint = entryPoint, Sequence = sequence, Phase = phase, Stage = stage, Thread = thread,
                    Utc = DateTime.UtcNow.ToString("o"), ElapsedTicks = session.Clock.ElapsedTicks,
                    Error = phase == "Fault" ? (error == null ? "Caught exception; managed detail unavailable." : error.ToString()) : null,
                    FaultCaught = phase == "Fault" ? (bool?)caught : null,
                    ErrorIsComException = phase == "Fault" ? (bool?)(error is COMException) : null }, stage == null && phase != "Entry");
            }
            catch (Exception publication) { Fail(publication); }
        }

        /// <summary>Accepts only the bounded canonical serializer format; duplicated properties cannot survive canonical equality.</summary>
        internal static Identity DecodeRequest(string text, string nonce)
        {
            Guid guid;
            if (!Guid.TryParseExact(nonce, "N", out guid) || text == null || Encoding.UTF8.GetByteCount(text) > MaximumBytes)
                throw new InvalidOperationException("A bounded owned shutdown request is required.");
            var json = new JavaScriptSerializer();
            var fields = json.DeserializeObject(text) as IDictionary<string, object>;
            if (fields == null || fields.Count != 3 || !fields.ContainsKey("Version") || !(fields["Version"] is int) || (int)fields["Version"] != 1 ||
                !fields.ContainsKey("Nonce") || !string.Equals(fields["Nonce"] as string, nonce, StringComparison.Ordinal) || !fields.ContainsKey("Identity") ||
                !string.Equals(json.Serialize(fields), text, StringComparison.Ordinal))
                throw new InvalidOperationException("The version, nonce and canonical request fields must be exact.");
            return DecodeIdentity(fields["Identity"]);
        }

        /// <summary>Requires exact identity members without permissive number or string coercions.</summary>
        internal static Identity DecodeIdentity(object value)
        {
            var json = new JavaScriptSerializer();
            var identity = value as IDictionary<string, object>;
            string[] names = { "ProcessId", "ProcessStartedUtc", "HostImagePath", "ProductPath", "ProductMvid", "ProductSha256", "ThreadId" };
            if (identity == null || identity.Count != names.Length || names.Any(name => !identity.ContainsKey(name)) ||
                !(identity["ProcessId"] is int) || (!(identity["ThreadId"] is int) && !(identity["ThreadId"] is long)) ||
                names.Where(name => name != "ProcessId" && name != "ThreadId").Any(name => !(identity[name] is string)))
                throw new InvalidOperationException("Shutdown identity fields and their types must be exact.");
            return json.Deserialize<Identity>(json.Serialize(identity));
        }

        /// <summary>Validates requested syntax and compares it to independent process/product observations.</summary>
        internal static void RequireSameIdentity(Identity expected, Identity actual)
        {
            DateTime birth; Guid mvid;
            if (expected == null || actual == null || expected.ProcessId <= 0 || expected.ThreadId == 0 ||
                !DateTime.TryParseExact(expected.ProcessStartedUtc, "o", CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out birth) || birth.Kind != DateTimeKind.Utc ||
                !Guid.TryParseExact(expected.ProductMvid, "D", out mvid) || !IsLocalCanonical(expected.HostImagePath) || !IsLocalCanonical(expected.ProductPath) ||
                expected.ProductSha256 == null || expected.ProductSha256.Length != 64 || expected.ProductSha256.Any(value => !Uri.IsHexDigit(value)) ||
                expected.ProcessId != actual.ProcessId || expected.ProcessStartedUtc != actual.ProcessStartedUtc || expected.ThreadId != actual.ThreadId ||
                !string.Equals(expected.HostImagePath, actual.HostImagePath, StringComparison.OrdinalIgnoreCase) ||
                !string.Equals(expected.ProductPath, actual.ProductPath, StringComparison.OrdinalIgnoreCase) ||
                !string.Equals(expected.ProductMvid, actual.ProductMvid, StringComparison.OrdinalIgnoreCase) ||
                !string.Equals(expected.ProductSha256, actual.ProductSha256, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The owned process generation, native thread and actually loaded product must match.");
        }

        private static bool IsLocalCanonical(string path)
        { return !string.IsNullOrWhiteSpace(path) && Path.IsPathRooted(path) && !path.StartsWith(@"\\", StringComparison.Ordinal) &&
            path.Length >= 3 && path[1] == ':' && path[2] == '\\' && string.Equals(path, Path.GetFullPath(path), StringComparison.OrdinalIgnoreCase); }

        /// <summary>Bounds ordinary paths to a fixed temporary GUID child; this is not a hostile-race filesystem security primitive.</summary>
        internal static string RequireRoot(string root, string temporaryRoot)
        {
            Guid guid;
            if (!IsLocalCanonical(root) || !Guid.TryParseExact(Path.GetFileName(root), "N", out guid) ||
                !string.Equals(Path.GetDirectoryName(root), Path.Combine(Path.GetFullPath(temporaryRoot), DirectoryName).TrimEnd('\\'), StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("An owned local temporary shutdown diagnostic root is required.");
            RequireNoReparse(root);
            if (!Directory.Exists(root)) throw new DirectoryNotFoundException("The shutdown root must be prepared before launch.");
            return root;
        }

        /// <summary>Rejects existing linked components, including a linked request file, before reads or publication.</summary>
        internal static void RequireNoReparse(string path)
        { RequireNoReparse(path, File.GetAttributes); }

        /// <summary>Pure path-component observation seam; production reads actual Windows attributes.</summary>
        internal static void RequireNoReparse(string path, Func<string, FileAttributes> readAttributes)
        {
            if (readAttributes == null) throw new ArgumentNullException(nameof(readAttributes));
            string current = Path.GetFullPath(path);
            while (!string.IsNullOrEmpty(current))
            {
                try { if ((readAttributes(current) & FileAttributes.ReparsePoint) != 0) throw new InvalidOperationException("Linked diagnostic paths are refused."); }
                catch (FileNotFoundException) { }
                catch (DirectoryNotFoundException) { }
                current = Path.GetDirectoryName(current);
            }
        }

        /// <summary>Writes closed CreateNew files, moves once, then publishes terminal readiness; never overwrites or retries.</summary>
        internal static void Publish(string root, int sequence, object value, bool terminal)
        {
            if (sequence < 1 || sequence > MaximumEvents) throw new ArgumentOutOfRangeException(nameof(sequence));
            RequireNoReparse(root);
            string path = Path.Combine(root, sequence.ToString("D6", CultureInfo.InvariantCulture) + ".json"), temporary = path + ".tmp";
            RequireNoReparse(path); RequireNoReparse(temporary); RequireNoReparse(path + ".ready");
            string text = new JavaScriptSerializer().Serialize(value);
            if (Encoding.UTF8.GetByteCount(text) > MaximumBytes) throw new InvalidOperationException("Shutdown event exceeds its bound.");
            using (var file = new FileStream(temporary, FileMode.CreateNew, FileAccess.Write, FileShare.None))
            using (var writer = new StreamWriter(file, new UTF8Encoding(false))) writer.Write(text);
            File.Move(temporary, path);
            if (terminal) using (var ready = new FileStream(path + ".ready", FileMode.CreateNew, FileAccess.Write, FileShare.None)) { }
        }

        private static Identity CaptureIdentity()
        {
            using (var process = Process.GetCurrentProcess())
            {
                if (!string.Equals(process.ProcessName, "EXCEL", StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("This observation is restricted to the explicitly owned Excel qualification.");
                string product = typeof(AddIn).Assembly.Location, hash;
                using (var file = new FileStream(product, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var sha = SHA256.Create()) hash = BitConverter.ToString(sha.ComputeHash(file)).Replace("-", "");
                return new Identity { ProcessId = process.Id, ProcessStartedUtc = process.StartTime.ToUniversalTime().ToString("o"),
                    HostImagePath = process.MainModule.FileName, ProductPath = product,
                    ProductMvid = typeof(AddIn).Module.ModuleVersionId.ToString("D"), ProductSha256 = hash, ThreadId = GetCurrentThreadId() };
            }
        }

        private static ThreadIdentity CaptureThread()
        { return new ThreadIdentity { ManagedThreadId = Thread.CurrentThread.ManagedThreadId, NativeThreadId = GetCurrentThreadId(), Apartment = Thread.CurrentThread.GetApartmentState().ToString() }; }
        [DllImport("kernel32.dll")] private static extern uint GetCurrentThreadId();
    }
}
