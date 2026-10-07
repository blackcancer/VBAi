using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;

namespace VBAi.Tests.Integration
{
    /// <summary>Opt-in read-only CLR capture on one already-owned private-desktop Excel; no function evaluation or replay.</summary>
    internal sealed class Q026OptionsGuardTrace : IDisposable
    {
        [DllImport("kernel32.dll", SetLastError = true)] private static extern bool DebugBreakProcess(IntPtr process);
        [DllImport("kernel32.dll", SetLastError = true)] private static extern bool CheckRemoteDebuggerPresent(IntPtr process, out bool present);
        [DllImport("kernel32.dll")] private static extern uint GetACP();
        private Process target, debugger;
        private Task<string> stdout, stderr;
        private string directory, log;
        private bool stopped;
        private readonly Dictionary<string, object> record = new Dictionary<string, object>();

        internal static bool NeedsGuardWarmupIfRequested()
        {
            string path = Environment.GetEnvironmentVariable("VBAi_TEST_Q026_CLR_TRACE_PLAN");
            if (string.IsNullOrEmpty(path)) return false;
            var plan = LoadValidatedPlan(path);
            return plan.TryGetValue("TraceMode", out object mode) && Equals(mode, "ExactGuardILBreakpoint");
        }

        private static Dictionary<string, object> LoadValidatedPlan(string planPath)
        {
            if (!Path.IsPathRooted(planPath) || new FileInfo(planPath).Length > 65536) throw new InvalidOperationException("A bounded absolute CLR trace plan is required.");
            var json = new JavaScriptSerializer();
            var plan = json.Deserialize<Dictionary<string, object>>(File.ReadAllText(planPath));
            var preflight = json.Deserialize<Dictionary<string, object>>(File.ReadAllText((string)plan["Preflight"]));
            string cdb = (string)plan["CdbPath"], script = (string)plan["TraceScript"];
            if (!Equals(preflight["State"], "SYNTHETIC_PREFLIGHT_VERIFIED") || !Equals(preflight["FullSyntheticTabsEqual"], true) ||
                !Equals(preflight["DebuggerDetachedVerified"], true) || !Equals(preflight["OriginalHandleExitObserved"], true) ||
                Convert.ToInt32(preflight["TargetExitCode"]) != 0 || Convert.ToInt32(preflight["DebuggerExitCode"]) != 0 ||
                Hash(cdb) != (string)preflight["CdbSha256"] || Hash(script) != (string)preflight["TraceScriptSha256"] ||
                Hash(cdb) != (string)plan["CdbSha256"] || Hash(script) != (string)plan["TraceScriptSha256"] ||
                typeof(VbeSession).Module.ModuleVersionId.ToString("D") != (string)plan["ProductMvid"])
                throw new InvalidOperationException("CLR preflight, frozen collector or candidate differs; no attachment.");
            if (plan.TryGetValue("TraceMode", out object mode) && !Equals(mode, "FirstChanceClrException"))
            {
                if (!Equals(mode, "ExactGuardILBreakpoint") || !Equals(preflight["TraceMode"], mode) ||
                    !Equals(preflight["ExactGuardILCaptured"], true) || !Equals(preflight["ProductMvid"], plan["ProductMvid"]) ||
                    !Equals(preflight["GuardILOffset"], plan["GuardILOffset"]) ||
                    !Equals(preflight["ProductSha256"], Hash(typeof(VbeSession).Assembly.Location)))
                    throw new InvalidOperationException("Exact guard breakpoint mode lacks its own complete frozen preflight.");
                int offset = Convert.ToInt32(plan["GuardILOffset"]);
                var method = typeof(VbeSession).Assembly.GetType("VBAi.VbeDebugWindows", true)
                    .GetMethods(BindingFlags.Static | BindingFlags.NonPublic)
                    .Single(item => item.Name == "SetVbeOption" && item.GetParameters().Length == 2);
                byte[] il = method.GetMethodBody().GetILAsByteArray();
                if (offset < 0 || offset > il.Length - 5 || il[offset] != 0x72 ||
                    method.Module.ResolveString(BitConverter.ToInt32(il, offset + 1)) != "VBE options changed since inspection; read them again.")
                    throw new InvalidOperationException("Frozen IL offset does not identify the exact stale-revision branch.");
            }
            return plan;
        }

        internal static Q026OptionsGuardTrace StartIfRequested(ExcelVbeFixture host, DateTime startUtc, string evidence)
        {
            string planPath = Environment.GetEnvironmentVariable("VBAi_TEST_Q026_CLR_TRACE_PLAN");
            if (string.IsNullOrEmpty(planPath)) return null;
            var trace = new Q026OptionsGuardTrace { directory = Path.Combine(evidence, "clr-guard") };
            Directory.CreateDirectory(trace.directory);
            try { trace.Start(host, startUtc, planPath); return trace; }
            catch (Exception primary)
            {
                trace.record["StartError"] = primary.ToString();
                try { trace.Dispose(); }
                catch (Exception detach) { throw new AggregateException("CLR trace activation and detachment both failed; host must be retained.", primary, detach); }
                throw;
            }
        }

        private void Start(ExcelVbeFixture host, DateTime startUtc, string planPath)
        {
            if (!Path.IsPathRooted(planPath) || new FileInfo(planPath).Length > 65536) throw new InvalidOperationException("A bounded absolute CLR trace plan is required.");
            string desktop = Environment.GetEnvironmentVariable("VBAi_TEST_DESKTOP_NAME");
            if (string.IsNullOrEmpty(desktop)) throw new InvalidOperationException("No foreground debugger fallback is allowed.");
            IsolatedTestDesktop.RequireCurrent(desktop);
            var plan = LoadValidatedPlan(planPath);
            string cdb = (string)plan["CdbPath"], script = (string)plan["TraceScript"], preflightPath = (string)plan["Preflight"];
            bool exactGuard = plan.TryGetValue("TraceMode", out object mode) && Equals(mode, "ExactGuardILBreakpoint");
            target = Process.GetProcessById(host.ProcessId);
            if (target.HasExited || target.StartTime.ToUniversalTime() != startUtc) throw new InvalidOperationException("Owned Excel start identity differs.");
            bool present;
            if (!CheckRemoteDebuggerPresent(target.Handle, out present) || present) throw new InvalidOperationException("A debugger is already present or cannot be observed.");
            record["ProcessId"] = host.ProcessId; record["ProcessStartUtc"] = startUtc.ToString("o");
            record["ProductMvid"] = plan["ProductMvid"]; record["Desktop"] = desktop; record["FixtureRoot"] = host.Root;
            record["NativePreferenceWrites"] = 0; record["FunctionEvaluations"] = 0; record["ForcedTerminations"] = 0;
            record["Preflight"] = preflightPath; record["State"] = "ATTACHING"; Save();
            record["TraceMode"] = exactGuard ? "ExactGuardILBreakpoint" : "FirstChanceClrException";
            log = Path.Combine(directory, "trace.cdb.log");
            string provider = Path.Combine(Path.GetDirectoryName(cdb), "winext", "JsProvider.dll");
            if (!File.Exists(provider)) throw new InvalidOperationException("Installed JavaScript debugger provider is absent.");
            string commands = Path.Combine(directory, "commands.txt");
            string content = ".logopen /u \"" + Forward(log) + "\"\n.loadby sos clr\n.load \"" + Forward(provider) +
                "\"\n.scriptload \"" + Forward(script) + "\"\ndx @$scriptContents.configure(" + host.ProcessId +
                ")\n" + (exactGuard ? "sxd clr\ndx @$scriptContents.armGuardBreakpoint(" + Convert.ToInt32(plan["GuardILOffset"]) + ")\n" :
                    "sxe -c \"dx @$scriptContents.capture();gn\" clr\n") + "g\n";
            File.WriteAllText(commands, content, Encoding.GetEncoding(checked((int)GetACP()), EncoderFallback.ExceptionFallback, DecoderFallback.ExceptionFallback));
            var info = new ProcessStartInfo(cdb, "-pd -p " + host.ProcessId + " -netsyms:no -cf \"" + commands + "\"")
            {
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardInput = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            };
            debugger = Process.Start(info); stdout = debugger.StandardOutput.ReadToEndAsync(); stderr = debugger.StandardError.ReadToEndAsync();
            var watch = Stopwatch.StartNew();
            while (!debugger.HasExited && watch.Elapsed.TotalSeconds < 10)
            {
                string output = ReadLog();
                if (output.Contains("Q026_OPTIONS_TRACE_READY pid=" + host.ProcessId) &&
                    (!exactGuard || output.Contains("Q026_GUARD_IL_ARMED offset=" + Convert.ToInt32(plan["GuardILOffset"]) + " breakpoint=")) &&
                    CheckRemoteDebuggerPresent(target.Handle, out present) && present)
                {
                    record["State"] = "ARMED"; record["DebuggerProcessId"] = debugger.Id; Save(); return;
                }
                Thread.Sleep(50);
            }
            throw new InvalidOperationException("CLR collector did not arm within its deadline; no preference scenario may run.");
        }

        public void Dispose()
        {
            if (stopped) return;
            stopped = true; // A failed detach is uncertain; never issue the break/detach again.
            try
            {
                if (debugger != null && !debugger.HasExited)
                {
                    if (target.HasExited)
                    {
                        if (!debugger.WaitForExit(10000)) throw new InvalidOperationException("Target exited but debugger did not exit; no forced cleanup.");
                    }
                    else
                    {
                        debugger.StandardInput.WriteLine("sxd clr"); debugger.StandardInput.WriteLine("bc *");
                        debugger.StandardInput.WriteLine(".logclose"); debugger.StandardInput.WriteLine("qd"); debugger.StandardInput.Flush();
                        if (!DebugBreakProcess(target.Handle) || !debugger.WaitForExit(10000))
                            throw new InvalidOperationException("One owned debugger detach could not be verified; retain the host without replay.");
                    }
                }
                if (debugger != null)
                {
                    record["DebuggerExitCode"] = debugger.ExitCode;
                    File.WriteAllText(Path.Combine(directory, "cdb.stdout.log"), stdout.GetAwaiter().GetResult());
                    File.WriteAllText(Path.Combine(directory, "cdb.stderr.log"), stderr.GetAwaiter().GetResult());
                    if (debugger.ExitCode != 0) throw new InvalidOperationException("Debugger exited abnormally; retain the host.");
                }
                if (target != null && !target.HasExited)
                {
                    bool present;
                    if (!CheckRemoteDebuggerPresent(target.Handle, out present) || present) throw new InvalidOperationException("Debugger detachment is not independently verified.");
                }
                record["DebuggerDetachedVerified"] = true; record["State"] = "DETACHED";
            }
            catch (Exception error) { record["State"] = "DETACH_UNVERIFIED"; record["Error"] = error.ToString(); throw; }
            finally { Save(); }
        }

        private string ReadLog() { if (log == null || !File.Exists(log)) return ""; using (var stream = new FileStream(log, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)) using (var reader = new StreamReader(stream, Encoding.Unicode)) return reader.ReadToEnd(); }
        private static string Hash(string path) { using (var file = File.OpenRead(path)) using (var sha = SHA256.Create()) return BitConverter.ToString(sha.ComputeHash(file)).Replace("-", ""); }
        private static string Forward(string path) { if (!Path.IsPathRooted(path) || path.IndexOf('"') >= 0 || path.IndexOf('\n') >= 0 || path.IndexOf('\r') >= 0) throw new InvalidOperationException("Unsafe debugger path."); return path.Replace('\\', '/'); }
        private void Save() { File.WriteAllText(Path.Combine(directory, "controller.json"), new JavaScriptSerializer().Serialize(record)); }
    }
}
