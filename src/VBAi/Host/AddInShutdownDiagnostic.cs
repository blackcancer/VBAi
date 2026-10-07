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

        /// <summary>Opt-in environment variable naming the prepared temporary diagnostic root.</summary>
        internal const string EnvironmentName = "VBAi_TEST_ADDIN_SHUTDOWN_ROOT";

        /// <summary>Fixed child directory required beneath the local temporary root.</summary>
        internal const string DirectoryName = "VBAi.AddInShutdownDiagnostic";

        /// <summary>Hard caps for nested cleanup invocations, published events, and serialized bytes per event.</summary>
        internal const int MaximumInvocations = 16, MaximumEvents = 2048, MaximumBytes = 16384;

        /// <summary>The actual process generation, loaded file and expected native callback thread; contains no project lease.</summary>
        internal sealed class Identity
        {

            /// <summary>Gets or sets the process id.</summary>
            /// <value>Current process id exposed by identity.</value>
            public int ProcessId { get; set; }

            /// <summary>Gets or sets the process started utc.</summary>
            /// <value>Current process started utc exposed by identity.</value>
            public string ProcessStartedUtc { get; set; }

            /// <summary>Gets or sets the host image path.</summary>
            /// <value>Current host image path exposed by identity.</value>
            public string HostImagePath { get; set; }

            /// <summary>Gets or sets the product path.</summary>
            /// <value>Current product path exposed by identity.</value>
            public string ProductPath { get; set; }

            /// <summary>Gets or sets the product mvid.</summary>
            /// <value>Current product mvid exposed by identity.</value>
            public string ProductMvid { get; set; }

            /// <summary>Gets or sets the product sha256.</summary>
            /// <value>Current product sha256 exposed by identity.</value>
            public string ProductSha256 { get; set; }

            /// <summary>Gets or sets the thread id.</summary>
            /// <value>Current thread id exposed by identity.</value>
            public uint ThreadId { get; set; }
        }

        /// <summary>Caller values actually observed at an individual marker, independently of a requested identity.</summary>
        internal sealed class ThreadIdentity
        {

            /// <summary>Gets or sets the managed thread id.</summary>
            /// <value>Current managed thread id exposed by thread identity.</value>
            public int ManagedThreadId { get; set; }

            /// <summary>Gets or sets the native thread id.</summary>
            /// <value>Current native thread id exposed by thread identity.</value>
            public uint NativeThreadId { get; set; }

            /// <summary>Gets or sets the apartment.</summary>
            /// <value>Current apartment exposed by thread identity.</value>
            public string Apartment { get; set; }
        }

        /// <summary>Bounded per-request publication state; no delegates capture an add-in, window or COM object.</summary>
        private sealed class Session
        {

            /// <summary>Nonce read from the exact admitted request and repeated in every published receipt.</summary>
            internal string Nonce;

            /// <summary>Independently observed process, loaded product, and owner-thread identity bound to the request.</summary>
            internal Identity Identity;

            /// <summary>Fresh native/managed thread observation invoked for every event publication.</summary>
            internal Func<ThreadIdentity> ObserveThread;

            /// <summary>Bounded receipt writer receiving event sequence, payload, and terminal-marker state.</summary>
            internal Action<int, object, bool> Publish;

            /// <summary>Monotonic clock used to report elapsed cleanup ticks.</summary>
            internal readonly Stopwatch Clock = Stopwatch.StartNew();

            /// <summary>Counts admitted nested invocations and monotonically assigned events for this request.</summary>
            internal int Invocations, Sequence;

            /// <summary>First observer or publication failure; once set, later events are suppressed.</summary>
            internal Exception Error;

            /// <summary>Nested cleanup invocations still active, ordered outermost to innermost.</summary>
            internal readonly List<AddInShutdownDiagnostic> Active = new List<AddInShutdownDiagnostic>();
        }

        /// <summary>Serializes admission of the process-wide shutdown request.</summary>
        private static readonly object RuntimeLock = new object();

        /// <summary>Admitted process-wide observer session shared by nested shutdown callbacks.</summary>
        private static Session runtimeSession;

        /// <summary>Canonical request JSON retained only while its admitted runtime session is active.</summary>
        private static string runtimeRequest;

        /// <summary>Shared bounded publication state for this invocation and its nested children.</summary>
        private readonly Session session;

        /// <summary>Original add-in shutdown entry point and exact active parent invocation ID.</summary>
        private readonly string entryPoint, parent;

        /// <summary>Currently entered cleanup stage awaiting its matching return or caught-fault marker.</summary>
        private string activeStage;

        /// <summary>Whether this invocation has published its single terminal result.</summary>
        private bool terminal;

        /// <summary>Gets or sets the invocation id.</summary>
        /// <value>Current invocation id exposed by add in shutdown diagnostic.</value>
        internal string InvocationId { get; private set; }

        /// <summary>Gets the diagnostic failed.</summary>
        /// <value>Current diagnostic failed exposed by add in shutdown diagnostic.</value>
        internal bool DiagnosticFailed { get { return session.Error != null; } }

        /// <summary>Gets the publication error.</summary>
        /// <value>Current publication error exposed by add in shutdown diagnostic.</value>
        internal Exception PublicationError { get { return session.Error; } }

        /// <summary>Creates one linked observation for an existing shutdown cleanup invocation.</summary>
        /// <param name="shared">Request-scoped session shared with nested cleanup callbacks.</param>
        /// <param name="entry">Allowed existing entry point: OnBeginShutdown, OnDisconnection, or Dispose.</param>
        /// <param name="parentId">Exact active parent invocation ID, or null for a top-level callback.</param>
        private AddInShutdownDiagnostic(Session shared, string entry, string parentId)
        { session = shared; entryPoint = entry; parent = parentId; InvocationId = Guid.NewGuid().ToString("N"); }

        /// <summary>Returns immediately when disabled; an admission or publication error never prevents original cleanup.</summary>
        /// <param name="entryPoint">Existing shutdown callback name to observe.</param>
        /// <returns>An admitted observer, or <see langword="null"/> when disabled or any opt-in admission check fails.</returns>
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
        /// <param name="json">Canonical bounded request JSON published by the owned fixture.</param>
        /// <param name="nonce">Expected 32-character GUID nonce from the request directory name.</param>
        /// <param name="entryPoint">Existing shutdown callback being observed.</param>
        /// <param name="observe">Independent process/product identity reader.</param>
        /// <param name="observeThread">Reader called for every event to verify the expected STA.</param>
        /// <param name="publish">Receipt writer receiving sequence, payload, and terminal state.</param>
        /// <returns>A new invocation after request and live identity match exactly.</returns>
        internal static AddInShutdownDiagnostic Begin(string json, string nonce, string entryPoint,
            Func<Identity> observe, Func<ThreadIdentity> observeThread, Action<int, object, bool> publish)
        {
            if (observe == null || observeThread == null || publish == null) throw new ArgumentNullException("Shutdown diagnostic dependencies");
            Identity expected = DecodeRequest(json, nonce), actual = observe();
            RequireSameIdentity(expected, actual);
            return NewInvocation(CreateSession(nonce, actual, observeThread, publish), entryPoint, null);
        }

        /// <summary>Copies the admitted identity into request-scoped state and binds thread observation and publication.</summary>
        /// <param name="nonce">Admitted request nonce.</param><param name="actual">Independently observed process and product identity.</param>
        /// <param name="observe">Fresh owning-thread reader.</param><param name="publish">Bounded event writer.</param>
        /// <returns>New session with a monotonic elapsed clock and empty invocation stack.</returns>
        private static Session CreateSession(string nonce, Identity actual, Func<ThreadIdentity> observe,
            Action<int, object, bool> publish)
        {
            var json = new JavaScriptSerializer();
            return new Session
            {
                Nonce = nonce,
                Identity = json.Deserialize<Identity>(json.Serialize(actual)),
                ObserveThread = observe,
                Publish = publish
            };
        }

        /// <summary>Admits a bounded nested cleanup invocation only beneath its exact currently active parent.</summary>
        /// <param name="shared">Request-scoped session whose invocation stack and limits are updated.</param>
        /// <param name="entry">Existing shutdown entry point to record.</param><param name="parent">Active parent invocation ID, or null for a top-level invocation.</param>
        /// <returns>The pushed invocation, or <see langword="null"/> after recording the first diagnostic failure.</returns>
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
        /// <param name="childEntryPoint">Existing cleanup entry point invoked beneath this active operation.</param>
        /// <returns>A linked child invocation, or <see langword="null"/> if admission fails.</returns>
        internal AddInShutdownDiagnostic Child(string childEntryPoint)
        { return NewInvocation(session, childEntryPoint, InvocationId); }

        /// <summary>Marks an operation about to run; no native or managed resource is captured.</summary>
        /// <param name="stage">Non-empty label for the existing cleanup operation about to run.</param>
        internal void Enter(string stage)
        {
            if (DiagnosticFailed) return;
            if (terminal || activeStage != null || string.IsNullOrWhiteSpace(stage)) { Fail(new InvalidOperationException("Invalid shutdown stage entry.")); return; }
            activeStage = stage; Record("Entry", stage, null);
        }

        /// <summary>Marks only the normal return of the previously entered existing operation.</summary>
        /// <param name="stage">Exact label passed to the preceding <see cref="Enter"/> call.</param>
        internal void Returned(string stage)
        {
            if (DiagnosticFailed) return;
            if (terminal || activeStage != stage) { Fail(new InvalidOperationException("Shutdown return has no exact entered stage.")); return; }
            Record("Returned", stage, null); activeStage = null;
        }

        /// <summary>Records an error absorbed by an existing catch without turning it into a normal operation return.</summary>
        /// <param name="stage">Exact label for the existing stage whose catch absorbed the exception.</param>
        /// <param name="error">Caught exception recorded as a fault rather than a normal return.</param>
        internal void CaughtFault(string stage, Exception error)
        {
            if (DiagnosticFailed) return;
            if (terminal || activeStage != stage) { Fail(new InvalidOperationException("Shutdown caught fault has no exact stage.")); return; }
            Record("Fault", stage, error, true); activeStage = null;
        }

        /// <summary>Observes an existing null assignment, without claiming RCW release or native Close.</summary>
        /// <param name="field">Name of the managed reference field observed being assigned null; this does not imply COM release.</param>
        internal void ManagedReferenceCleared(string field)
        {
            if (DiagnosticFailed) return;
            if (terminal || activeStage != null || string.IsNullOrWhiteSpace(field)) { Fail(new InvalidOperationException("Invalid cleared-reference observation.")); return; }
            Record("ManagedReferenceCleared", field, null);
        }

        /// <summary>Records propagated failure; the caller rethrows the same original exception.</summary>
        /// <param name="error">Exception describing the error failure.</param>
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

        /// <summary>Publishes this invocation's terminal phase and removes it from the active nesting stack.</summary>
        /// <param name="phase">Terminal phase, either Returned or Fault.</param><param name="error">Optional exception included only for a fault phase.</param>
        private void Finish(string phase, Exception error)
        {
            if (session.Active.Count == 0 || session.Active.Last() != this) { Fail(new InvalidOperationException("Shutdown invocation cannot finish ahead of its active child.")); return; }
            Record(phase, null, error); terminal = true;
            session.Active.RemoveAt(session.Active.Count - 1);
        }

        /// <summary>Stores the first diagnostic-only failure and suppresses subsequent publication.</summary>
        /// <param name="error">Admission, invariant, observation, or receipt-publication failure.</param>
        private void Fail(Exception error) { session.Error = session.Error ?? error; }

        /// <summary>Stops all subsequent publication on the first diagnostic error; never changes cleanup propagation.</summary>
        /// <param name="phase">Lifecycle phase such as Entry, Returned, ManagedReferenceCleared, or Fault.</param>
        /// <param name="stage">Optional label for an individual existing cleanup operation.</param>
        /// <param name="error">Exception associated with a fault event; its message may appear in the diagnostic receipt.</param>
        /// <param name="caught"><see langword="true"/> when the original cleanup caught and absorbed the fault.</param>
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
                session.Publish(sequence, new
                {
                    Version = 1,
                    session.Nonce,
                    session.Identity,
                    InvocationId,
                    ParentInvocationId = parent,
                    EntryPoint = entryPoint,
                    Sequence = sequence,
                    Phase = phase,
                    Stage = stage,
                    Thread = thread,
                    Utc = DateTime.UtcNow.ToString("o"),
                    session.Clock.ElapsedTicks,
                    Error = phase == "Fault" ? (error == null ? "Caught exception; managed detail unavailable." : error.ToString()) : null,
                    FaultCaught = phase == "Fault" ? (bool?)caught : null,
                    ErrorIsComException = phase == "Fault" ? (bool?)(error is COMException) : null
                }, stage == null && phase != "Entry");
            }
            catch (Exception publication) { Fail(publication); }
        }

        /// <summary>Accepts only the bounded canonical serializer format; duplicated properties cannot survive canonical equality.</summary>
        /// <param name="text">UTF-8 request JSON, bounded by <see cref="MaximumBytes"/> and required to use the canonical serializer format.</param>
        /// <param name="nonce">Expected request nonce, encoded as a 32-character GUID without separators.</param>
        /// <returns>Exactly typed process, product, and native-thread identity from the request.</returns>
        /// <exception cref="InvalidOperationException">The nonce, size, version, fields, types, or canonical serialization do not match.</exception>
        internal static Identity DecodeRequest(string text, string nonce)
        {
            if (!Guid.TryParseExact(nonce, "N", out _) || text == null || Encoding.UTF8.GetByteCount(text) > MaximumBytes)
                throw new InvalidOperationException("A bounded owned shutdown request is required.");
            var json = new JavaScriptSerializer();
            if (!(json.DeserializeObject(text) is IDictionary<string, object> fields) || fields.Count != 3 || !fields.ContainsKey("Version") || !(fields["Version"] is int v) || v != 1 ||
                !fields.ContainsKey("Nonce") || !string.Equals(fields["Nonce"] as string, nonce, StringComparison.Ordinal) || !fields.ContainsKey("Identity") ||
                !string.Equals(json.Serialize(fields), text, StringComparison.Ordinal))
                throw new InvalidOperationException("The version, nonce and canonical request fields must be exact.");
            return DecodeIdentity(fields["Identity"]);
        }

        /// <summary>Requires exact identity members without permissive number or string coercions.</summary>
        /// <param name="value">Deserialized JSON object containing the seven identity fields.</param>
        /// <returns>An identity object after rejecting missing, extra, or coerced field values.</returns>
        internal static Identity DecodeIdentity(object value)
        {
            var json = new JavaScriptSerializer();
            string[] names = { "ProcessId", "ProcessStartedUtc", "HostImagePath", "ProductPath", "ProductMvid", "ProductSha256", "ThreadId" };
            if (!(value is IDictionary<string, object> identity) || identity.Count != names.Length || names.Any(name => !identity.ContainsKey(name)) ||
                !(identity["ProcessId"] is int) || (!(identity["ThreadId"] is int) && !(identity["ThreadId"] is long)) ||
                names.Where(name => name != "ProcessId" && name != "ThreadId").Any(name => !(identity[name] is string)))
                throw new InvalidOperationException("Shutdown identity fields and their types must be exact.");
            return json.Deserialize<Identity>(json.Serialize(identity));
        }

        /// <summary>Validates requested syntax and compares it to independent process/product observations.</summary>
        /// <param name="expected">Identity decoded from the owned shutdown request.</param>
        /// <param name="actual">Independently captured live process, loaded product, and owner-thread identity.</param>
        internal static void RequireSameIdentity(Identity expected, Identity actual)
        {
            if (expected == null || actual == null || expected.ProcessId <= 0 || expected.ThreadId == 0 ||
                !DateTime.TryParseExact(expected.ProcessStartedUtc, "o", CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out DateTime birth) || birth.Kind != DateTimeKind.Utc ||
                !Guid.TryParseExact(expected.ProductMvid, "D", out Guid mvid) || !IsLocalCanonical(expected.HostImagePath) || !IsLocalCanonical(expected.ProductPath) ||
                expected.ProductSha256 == null || expected.ProductSha256.Length != 64 || expected.ProductSha256.Any(value => !Uri.IsHexDigit(value)) ||
                expected.ProcessId != actual.ProcessId || expected.ProcessStartedUtc != actual.ProcessStartedUtc || expected.ThreadId != actual.ThreadId ||
                !string.Equals(expected.HostImagePath, actual.HostImagePath, StringComparison.OrdinalIgnoreCase) ||
                !string.Equals(expected.ProductPath, actual.ProductPath, StringComparison.OrdinalIgnoreCase) ||
                !string.Equals(expected.ProductMvid, actual.ProductMvid, StringComparison.OrdinalIgnoreCase) ||
                !string.Equals(expected.ProductSha256, actual.ProductSha256, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The owned process generation, native thread and actually loaded product must match.");
        }

        /// <summary>Accepts only fully qualified drive-letter paths already in their normalized local form.</summary>
        /// <param name="path">Path to validate; UNC paths and non-canonical spellings are rejected.</param>
        /// <returns><see langword="true"/> when the path is rooted on a drive and equals its full-path form.</returns>
        private static bool IsLocalCanonical(string path)
        {
            return !string.IsNullOrWhiteSpace(path) && Path.IsPathRooted(path) && !path.StartsWith(@"\\", StringComparison.Ordinal) &&
            path.Length >= 3 && path[1] == ':' && path[2] == '\\' && string.Equals(path, Path.GetFullPath(path), StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>Bounds ordinary paths to a fixed temporary GUID child; this is not a hostile-race filesystem security primitive.</summary>
        /// <param name="root">Prepared request directory expected to be a GUID child of the fixed diagnostic directory.</param>
        /// <param name="temporaryRoot">Local temporary root against which the diagnostic directory is resolved.</param>
        /// <returns>The validated canonical request directory.</returns>
        internal static string RequireRoot(string root, string temporaryRoot)
        {
            if (!IsLocalCanonical(root) || !Guid.TryParseExact(Path.GetFileName(root), "N", out _) ||
                !string.Equals(Path.GetDirectoryName(root), Path.Combine(Path.GetFullPath(temporaryRoot), DirectoryName).TrimEnd('\\'), StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("An owned local temporary shutdown diagnostic root is required.");
            RequireNoReparse(root);
            if (!Directory.Exists(root)) throw new DirectoryNotFoundException("The shutdown root must be prepared before launch.");
            return root;
        }

        /// <summary>Rejects existing linked components, including a linked request file, before reads or publication.</summary>
        /// <param name="path">File or directory path whose existing components must contain no reparse points.</param>
        internal static void RequireNoReparse(string path)
        { RequireNoReparse(path, File.GetAttributes); }

        /// <summary>Pure path-component observation seam; production reads actual Windows attributes.</summary>
        /// <param name="path">Path whose existing components are inspected up to the filesystem root.</param>
        /// <param name="readAttributes">Attribute reader; missing components are tolerated, but any observed reparse point is refused.</param>
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
        /// <param name="root">Validated per-request output directory.</param>
        /// <param name="sequence">One-based event sequence, bounded by <see cref="MaximumEvents"/>.</param>
        /// <param name="value">Serializable event payload; UTF-8 output must fit <see cref="MaximumBytes"/>.</param>
        /// <param name="terminal"><see langword="true"/> to create the one-shot ready marker after publishing the event file.</param>
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

        /// <summary>Captures the current Excel process generation, loaded VBAi image identity, hash, and native thread.</summary>
        /// <returns>Exact identity used to bind the request to this process and loaded assembly.</returns>
        /// <exception cref="InvalidOperationException">The current process is not Excel.</exception>
        private static Identity CaptureIdentity()
        {
            using (var process = Process.GetCurrentProcess())
            {
                if (!string.Equals(process.ProcessName, "EXCEL", StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("This observation is restricted to the explicitly owned Excel qualification.");
                string product = typeof(AddIn).Assembly.Location, hash;
                using (var file = new FileStream(product, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var sha = SHA256.Create()) hash = BitConverter.ToString(sha.ComputeHash(file)).Replace("-", "");
                return new Identity
                {
                    ProcessId = process.Id,
                    ProcessStartedUtc = process.StartTime.ToUniversalTime().ToString("o"),
                    HostImagePath = process.MainModule.FileName,
                    ProductPath = product,
                    ProductMvid = typeof(AddIn).Module.ModuleVersionId.ToString("D"),
                    ProductSha256 = hash,
                    ThreadId = GetCurrentThreadId()
                };
            }
        }

        /// <summary>Captures managed and native thread IDs and the apartment state at the observation point.</summary>
        /// <returns>Thread identity used to verify that every receipt is emitted from the expected STA.</returns>
        private static ThreadIdentity CaptureThread()
        { return new ThreadIdentity { ManagedThreadId = Thread.CurrentThread.ManagedThreadId, NativeThreadId = GetCurrentThreadId(), Apartment = Thread.CurrentThread.GetApartmentState().ToString() }; }

        /// <summary>Returns the native Windows identifier of the calling thread.</summary>
        /// <returns>Native thread ID from kernel32.</returns>
        [DllImport("kernel32.dll")] private static extern uint GetCurrentThreadId();
    }
}
