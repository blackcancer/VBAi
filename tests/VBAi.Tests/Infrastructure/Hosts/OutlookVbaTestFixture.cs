using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Web.Script.Serialization;
using Microsoft.Win32;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Integration
{
    /// <summary>Owns a fresh Outlook process and a previously absent OTM; never sends or saves an Outlook item.</summary>
    internal sealed class OutlookVbaTestFixture : IDisposable
    {
        private object application, item, inspector, commandBars;
        private Process process;
        private bool baselineAbsent, baselineVerified, disposed, runPending, hostTeardownRefused;
        private Exception startupFailure;
        private string referencesVersion;
        private IDictionary<string, object>[] baseline;
        private readonly Dictionary<string, string> ownedModules = new Dictionary<string, string>(StringComparer.Ordinal);
        private readonly List<object> evidence = new List<object>();
        private readonly OfficeCommandContainment commandContainment = new OfficeCommandContainment();
        // Retain every owned reference: final RCW release must not become implicit cleanup of unknown native work.
        private static readonly List<OutlookVbaTestFixture> retainedOutlookFixtures = new List<OutlookVbaTestFixture>();
        internal Func<int, object, IDictionary<string, object>> Dispatch = (pid, request) => VbeBridgeClient.Read(pid, request);
        internal Action VerifyOwnedProcess;
        internal OutlookVbaTestFixture() { VerifyOwnedProcess = RequireOnlyOwnedProcess; }
        internal string Root { get; private set; }
        internal int ProcessId { get; private set; }
        internal string Project { get; private set; }
        internal string OtmPath { get; private set; }

        [StructLayout(LayoutKind.Sequential)] private struct Disposition { public byte DeleteFile; }
        [StructLayout(LayoutKind.Sequential)] private struct AttributeTag { public uint Attributes, ReparseTag; }
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)] private static extern Microsoft.Win32.SafeHandles.SafeFileHandle CreateFile(
            string path, uint access, uint share, IntPtr security, uint creation, uint flags, IntPtr template);
        [DllImport("kernel32.dll", SetLastError = true)] private static extern bool GetFileInformationByHandleEx(
            Microsoft.Win32.SafeHandles.SafeFileHandle file, int informationClass, out AttributeTag attributes, uint size);
        [DllImport("kernel32.dll", SetLastError = true)] private static extern bool SetFileInformationByHandle(
            Microsoft.Win32.SafeHandles.SafeFileHandle file, int informationClass, ref Disposition disposition, uint size);

        internal static OutlookVbaTestFixture Start()
        {
            if (Environment.GetEnvironmentVariable("VBAi_RUN_OUTLOOK_TESTS") != "1")
                Assert.Inconclusive("Set VBAi_RUN_OUTLOOK_TESTS=1 for disposable Outlook VBA test qualification.");
            Assert.AreEqual(8, IntPtr.Size, "Registered Outlook qualification requires the x64 test host.");
            RequireNoOutlook();
            string otm = Path.GetFullPath(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Microsoft", "Outlook", "VbaProject.OTM"));
            if (File.Exists(otm) || Directory.Exists(otm))
                Assert.Inconclusive("BLOCKED: the personal Outlook VBA project exists. No project replacement or rename is permitted.");
            string parent = Path.GetDirectoryName(otm);
            if (Directory.Exists(parent) && (File.GetAttributes(parent) & FileAttributes.ReparsePoint) != 0)
                Assert.Inconclusive("BLOCKED: the Outlook VBA directory is a reparse point; fixed-file ownership is not established.");
            using (var profiles = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Office\16.0\Outlook\Profiles"))
                if (profiles == null || profiles.GetSubKeyNames().Length == 0)
                    Assert.Inconclusive("BLOCKED: Outlook has no configured profile. Qualification cannot configure one.");
            var type = Type.GetTypeFromProgID("Outlook.Application");
            if (type == null) Assert.Inconclusive("Classic Outlook is not installed.");
            var registration = RequireRegisteredCandidate();
            var fixture = new OutlookVbaTestFixture { OtmPath = otm, baselineAbsent = true };
            fixture.Root = Path.GetFullPath(Path.Combine(Environment.GetEnvironmentVariable("VBAi_OUTLOOK_RESULTS") ?? Path.Combine(Path.GetTempPath(), "VBAi-Outlook-tests"), Guid.NewGuid().ToString("N")));
            Directory.CreateDirectory(fixture.Root);
            Console.WriteLine("Outlook disposable qualification=" + fixture.Root);
            fixture.Save("registration-preflight.json", registration);
            try
            {
                RequireNoOutlook();
                Assert.IsFalse(File.Exists(otm), "An OTM appeared before activation; no host was activated.");
                fixture.application = Activator.CreateInstance(type);
                var launched = Process.GetProcessesByName("OUTLOOK");
                try
                {
                    Assert.AreEqual(1, launched.Length, "No unique new Outlook PID; no inspector or project mutation is permitted.");
                    fixture.ProcessId = launched[0].Id;
                    fixture.process = Process.GetProcessById(fixture.ProcessId);
                    _ = fixture.process.Handle;
                }
                finally { foreach (var current in launched) current.Dispose(); }
                fixture.item = ((dynamic)fixture.application).CreateItem(0);
                ((dynamic)fixture.item).Display(false);
                fixture.inspector = ((dynamic)fixture.item).GetInspector;
                fixture.commandBars = ((dynamic)fixture.inspector).CommandBars;
                ((dynamic)fixture.commandBars).ExecuteMso("VisualBasic");
                var status = fixture.Data("status");
                fixture.Save("loaded-identity.json", status);
                Assert.AreEqual(fixture.ProcessId, Convert.ToInt32(status["HostProcessId"]));
                Assert.AreEqual(typeof(VbeSession).Module.ModuleVersionId.ToString("D"), status["AssemblyModuleVersionId"]);
                Assert.AreEqual(HashFile(typeof(VbeSession).Assembly.Location), HashFile((string)status["AssemblyPath"]));
                var projects = fixture.Items("list_projects");
                Assert.AreEqual(1, projects.Length, "No unique new Outlook project; project mutation is refused.");
                string path = projects[0]["FileName"] as string;
                string name = projects[0]["Name"] as string;
                Assert.IsFalse(string.IsNullOrWhiteSpace(name), "The unique Outlook project must have a usable name.");
                string reportedPath = string.IsNullOrWhiteSpace(path) ? null : Path.GetFullPath(path);
                bool fixedOtmPath = string.Equals(otm, reportedPath, StringComparison.OrdinalIgnoreCase);
                if (reportedPath != null && !fixedOtmPath)
                    Assert.IsFalse(File.Exists(reportedPath) || Directory.Exists(reportedPath),
                        "The Outlook project reports an existing location outside the fixed absent-baseline OTM path; mutation is refused.");
                // An unpersisted Outlook project can report an opaque, nonexistent FileName such as 'Projet 1'.
                // Resolve that unique project by name; only the fixed OTM path is ever used for OTM recovery.
                fixture.Project = fixedOtmPath ? otm : name;
                fixture.baseline = fixture.ReadSources();
                Assert.AreEqual(1, fixture.baseline.Length, "Fresh Outlook must contain only its default session module.");
                Assert.AreEqual("ThisOutlookSession", fixture.baseline[0]["Name"]);
                string code = ((string)fixture.baseline[0]["Code"]).Replace("\r\n", "\n").Trim();
                Assert.IsTrue(code.Length == 0 || code.Equals("Option Explicit", StringComparison.OrdinalIgnoreCase),
                    "BLOCKED: the Outlook project already contains source. No mutation is permitted.");
                fixture.referencesVersion = (string)fixture.Data("list_references")["Version"];
                fixture.baselineVerified = true;
                fixture.Save("project-baseline.json", new { OtmPath = otm, InitiallyAbsent = true, Modules = fixture.baseline,
                    ReferencesVersion = fixture.referencesVersion, Project = fixture.Project, ProcessId = fixture.ProcessId,
                    ReportedFileName = path, SelectorKind = fixedOtmPath ? "FixedOtmPath" : "UniqueUnpersistedProjectName" });
                return fixture;
            }
            catch (Exception primary)
            {
                fixture.startupFailure = primary;
                try { fixture.Dispose(); }
                catch (Exception cleanup)
                {
                    try { fixture.Save("startup-failure.json", new { Primary = primary.ToString(), Cleanup = cleanup.ToString() }); }
                    catch (Exception recording)
                    { throw new AggregateException("Outlook fixture startup failed; cleanup and evidence recording also failed.", primary, cleanup, recording); }
                    throw new AggregateException("Outlook fixture startup failed; cleanup failure is reported separately.", primary, cleanup);
                }
                throw;
            }
        }

        internal IDictionary<string, object> Response(string command, params object[] pairs)
        {
            commandContainment.RequireTerminal();
            if (hostTeardownRefused) throw new InvalidOperationException("Outlook ownership is retained after uncertain native work; no request or cleanup retry is permitted.");
            Assert.IsNotNull(process, "An owned Outlook process is required.");
            Assert.IsFalse(process.HasExited, "Owned Outlook exited before dispatch.");
            VerifyOwnedProcess();
            var request = new Dictionary<string, object> { ["Command"] = command };
            if (Project != null) request["Project"] = Project;
            for (int index = 0; index < pairs.Length; index += 2) request[(string)pairs[index]] = pairs[index + 1];
            if (command == "run_vba_tests") runPending = true;
            var reply = commandContainment.Send(command, request, record => evidence.Add(record), FlushCommandEvidence,
                () => Dispatch(ProcessId, request), RetainUncertainOutlook);
            if (command == "run_vba_tests" && Equals(reply["Ok"], false) &&
                Convert.ToString(reply["Error"]).Contains("ExpectedProjectVersion")) runPending = false;
            if (command == "vba_test_run_status" && Equals(reply["Ok"], true))
            {
                var state = VbeBridgeClient.Object(reply["Data"]);
                if (Equals(state["Pending"], false))
                {
                    var report = state["Report"] as IDictionary<string, object>;
                    if (report != null) runPending = !Equals(report["uncertain"], false);
                }
            }
            return reply;
        }

        internal IDictionary<string, object> Data(string command, params object[] pairs)
        {
            var reply = Response(command, pairs);
            Assert.AreEqual(true, reply["Ok"], command + ": " + reply["Error"]);
            return VbeBridgeClient.Object(reply["Data"]);
        }

        internal IDictionary<string, object>[] Items(string command)
        {
            var reply = Response(command);
            Assert.AreEqual(true, reply["Ok"], command + ": " + reply["Error"]);
            return ((object[])reply["Data"]).Select(VbeBridgeClient.Object).ToArray();
        }

        internal void TrackOwnedModule(string name)
        {
            Assert.IsTrue(baselineVerified);
            Assert.IsFalse(baseline.Any(module => (string)module["Name"] == name));
            ownedModules[name] = (string)Data("read_module", "Module", name)["Sha256"];
        }

        private IDictionary<string, object>[] ReadSources()
        {
            return Items("list_modules").Select(module => {
                var source = Data("read_module", "Module", module["Name"]);
                return (IDictionary<string, object>)new Dictionary<string, object> {
                    ["Name"] = module["Name"], ["Code"] = source["Code"], ["Sha256"] = source["Sha256"] };
            }).OrderBy(module => (string)module["Name"], StringComparer.Ordinal).ToArray();
        }

        internal void Save(string name, object value)
        { File.WriteAllText(Path.Combine(Root, name), new JavaScriptSerializer { MaxJsonLength = 10 * 1024 * 1024 }.Serialize(value), new UTF8Encoding(false)); }

        public void Dispose()
        {
            if (disposed) return;
            disposed = true;
            var failures = new List<string>();
            bool restored = false, exitedNormally = false;
            if (runPending || commandContainment.Pending || commandContainment.Uncertain || hostTeardownRefused)
                RetainAndFail(failures, restored, exitedNormally, "Native execution or bridge delivery remains pending/uncertain; no Close, Quit, COM release, exit wait or OTM recovery was attempted.");
            if (process != null)
            {
                try
                {
                    if (baselineVerified && !runPending && !process.HasExited)
                    {
                        foreach (var module in ownedModules)
                        {
                            var source = Data("read_module", "Module", module.Key);
                            Assert.AreEqual(module.Value, source["Sha256"], "Owned source changed; cleanup must not remove concurrent user edits.");
                            var component = Data("component_properties", "Module", module.Key);
                            Data("remove_component", "Module", module.Key, "ExpectedComponentVersion", component["Version"],
                                "ExpectedProjectVersion", Data("project_properties")["Version"]);
                        }
                        var after = ReadSources();
                        Save("cleanup-sources.json", after);
                        Assert.AreEqual(baseline.Length, after.Length, "Unowned modules appeared; OTM retention is required.");
                        for (int index = 0; index < after.Length; index++)
                        {
                            Assert.AreEqual(baseline[index]["Name"], after[index]["Name"]);
                            Assert.AreEqual(baseline[index]["Sha256"], after[index]["Sha256"]);
                        }
                        Assert.AreEqual(referencesVersion, Data("list_references")["Version"]);
                        restored = true;
                    }
                    if (!restored) failures.Add("Project cleanup was not fully verified; any created OTM will be retained.");
                }
                catch (Exception error)
                {
                    failures.Add("Module cleanup: " + error.Message);
                    if (commandContainment.Pending || commandContainment.Uncertain || hostTeardownRefused)
                        RetainAndFail(failures, restored, exitedNormally, "Module cleanup delivery is uncertain; original ownership and any OTM are retained without another native command.");
                }
                try
                {
                    if (!process.HasExited)
                    {
                        VerifyOwnedProcess();
                        if (!runPending && inspector != null) ((dynamic)inspector).Close(1);
                        else if (runPending) failures.Add("Native run completion is unverified; inspector closure was refused and the owned process was retained.");
                    }
                }
                catch (Exception error)
                {
                    failures.Add("Unsaved inspector discard: " + error.Message);
                    RetainAndFail(failures, restored, exitedNormally, "Inspector closure failed with unknown native effect; no Quit or release was attempted.");
                }
                try
                {
                    if (process.HasExited)
                    {
                        exitedNormally = process.ExitCode == 0;
                        evidence.Add(new { Command = "Quit", Skipped = true, Reason = "Owned Outlook already exited.", ExitCode = process.ExitCode });
                    }
                    else
                    {
                        VerifyOwnedProcess();
                        if (!runPending && application != null) ((dynamic)application).Quit();
                        else failures.Add("Native run completion is unverified; Quit was refused and the owned process was retained.");
                    }
                }
                catch (Exception error)
                {
                    failures.Add("Normal Quit: " + error.Message);
                    RetainAndFail(failures, restored, exitedNormally, "Normal Quit failed with unknown native effect; no release, retry or OTM recovery was attempted.");
                }
            }
            foreach (var value in new[] { commandBars, inspector, item, application })
                if (value != null && Marshal.IsComObject(value)) try { Marshal.FinalReleaseComObject(value); }
                    catch (Exception error) { failures.Add("COM release: " + error.Message); }
            commandBars = inspector = item = application = null;
            GC.Collect(); GC.WaitForPendingFinalizers(); GC.Collect(); GC.WaitForPendingFinalizers();
            if (process != null)
            {
                try
                {
                    exitedNormally = process.WaitForExit(15000) && process.ExitCode == 0;
                    if (!exitedNormally) failures.Add("Owned Outlook did not exit normally; no process was killed and OTM was retained.");
                    if (exitedNormally && restored && baselineAbsent && failures.Count == 0) DeleteCreatedOtm();
                }
                catch (Exception error) { failures.Add("OTM recovery: " + error.Message); }
                finally { process.Dispose(); }
            }
            SaveLifecycle(failures, restored, exitedNormally);
            if (failures.Count != 0) Assert.Fail(string.Join("\n", failures));
        }

        private void FlushCommandEvidence()
        { Save("command-evidence.json", new { ProcessId, Project, runPending, commandContainment.Pending,
            commandContainment.Uncertain, commandContainment.Command, Steps = evidence }); }

        private void RetainUncertainOutlook()
        {
            hostTeardownRefused = true;
            lock (retainedOutlookFixtures)
                if (!retainedOutlookFixtures.Contains(this)) retainedOutlookFixtures.Add(this);
        }

        private void RetainAndFail(List<string> failures, bool restored, bool exitedNormally, string reason)
        {
            RetainUncertainOutlook();
            failures.Add(reason + " Retained PID=" + ProcessId);
            SaveLifecycle(failures, restored, exitedNormally);
            Assert.Fail(string.Join("\n", failures));
        }

        private void SaveLifecycle(List<string> failures, bool restored, bool exitedNormally)
        {
            Save("qualification-lifecycle.json", new { ProcessId, Project, OtmPath, InitiallyAbsent = baselineAbsent,
                ProjectRestored = restored, ExitedNormally = exitedNormally, StartupFailure = startupFailure?.ToString(),
                NativeExecutionUnsettled = runPending, BridgePending = commandContainment.Pending,
                BridgeUncertain = commandContainment.Uncertain, OwnershipRetained = hostTeardownRefused,
                ComReferencesRetained = hostTeardownRefused, ProcessHandleRetained = hostTeardownRefused && process != null,
                Failures = failures, Steps = evidence });
        }

        private void DeleteCreatedOtm()
        {
            RequireNoOutlook();
            if (!File.Exists(OtmPath)) { Save("otm-cleanup.json", new { Removed = false, Reason = "The baseline remains absent." }); return; }
            Assert.AreEqual(Path.GetFullPath(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Microsoft", "Outlook", "VbaProject.OTM")), OtmPath, true);
            Assert.IsFalse((File.GetAttributes(OtmPath) & FileAttributes.ReparsePoint) != 0);
            // Exclusive handle prevents writes/replacement; disposition deletes this exact opened file, not a later pathname occupant.
            // GENERIC_READ | DELETE, exclusive sharing, OPEN_EXISTING | FILE_FLAG_OPEN_REPARSE_POINT.
            using (var handle = CreateFile(OtmPath, 0x80010000, 0, IntPtr.Zero, 3, 0x00200000, IntPtr.Zero))
            {
                if (handle.IsInvalid) throw new System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error(), "Cannot lock the created OTM; deletion is refused.");
                AttributeTag attributes;
                if (!GetFileInformationByHandleEx(handle, 9, out attributes, 8))
                    throw new System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error(), "Cannot verify the created OTM handle.");
                Assert.AreEqual(0u, attributes.Attributes & (uint)FileAttributes.ReparsePoint, "OTM is a reparse point; deletion is refused.");
                using (var file = new FileStream(handle, FileAccess.Read))
                {
                    string hash;
                    using (var sha = SHA256.Create()) hash = BitConverter.ToString(sha.ComputeHash(file)).Replace("-", "");
                    file.Position = 0;
                    string backup = Path.Combine(Root, "created-VbaProject.OTM.backup");
                    using (var output = new FileStream(backup, FileMode.CreateNew, FileAccess.Write, FileShare.None))
                    { file.CopyTo(output); output.Flush(true); }
                    Assert.AreEqual(hash, HashFile(backup), "Created OTM backup did not match; deletion is refused.");
                    Save("otm-before-delete.json", new { OtmPath, Backup = backup, Sha256 = hash, Bytes = file.Length, ProcessId,
                        BaselineAbsent = true, ProjectRestored = true, OwnedProcessExitedNormally = true });
                    RequireNoOutlook();
                    var disposition = new Disposition { DeleteFile = 1 };
                    if (!SetFileInformationByHandle(file.SafeFileHandle, 4, ref disposition, 1))
                        throw new System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error(), "Exact-handle OTM deletion failed; backup retained.");
                }
            }
            Assert.IsFalse(File.Exists(OtmPath), "The created OTM remains; backup and evidence are retained.");
            Save("otm-cleanup.json", new { Removed = true, OtmPath });
        }

        private static object RequireRegisteredCandidate()
        {
            string hash = HashFile(typeof(VbeSession).Assembly.Location);
            var result = new List<object>();
            using (var classes = RegistryKey.OpenBaseKey(RegistryHive.ClassesRoot, RegistryView.Registry64))
                foreach (var entry in new[] {
                    new[] { "VBAi.AddIn", "{8E854243-087F-4D6C-9E0E-8622B0E50883}", "VBAi.AddIn" },
                    new[] { "VBAi.TestRuntime", "{5AF2F40B-939B-4CC6-A06C-F0C79841C031}", "VBAi.VbaTestRuntime" } })
                    using (var progId = classes.OpenSubKey(entry[0] + @"\CLSID"))
                    using (var server = classes.OpenSubKey(@"CLSID\" + entry[1] + @"\InprocServer32"))
                    {
                        Assert.IsNotNull(progId); Assert.IsNotNull(server, "Register the exact candidate before qualification; no fallback is permitted.");
                        Assert.AreEqual(entry[1], Convert.ToString(progId.GetValue("")), true);
                        Assert.AreEqual(entry[2], server.GetValue("Class"));
                        Assert.AreEqual(typeof(VbeSession).Assembly.FullName, server.GetValue("Assembly"));
                        Uri uri; Assert.IsTrue(Uri.TryCreate(Convert.ToString(server.GetValue("CodeBase")), UriKind.Absolute, out uri) && uri.IsFile);
                        Assert.AreEqual(hash, HashFile(uri.LocalPath), "Registered COM factory does not point at the exact candidate bytes.");
                        result.Add(new { ProgId = entry[0], Clsid = entry[1], AssemblyPath = uri.LocalPath, Sha256 = hash });
                    }
            return result;
        }

        private void RequireOnlyOwnedProcess()
        {
            var processes = Process.GetProcessesByName("OUTLOOK");
            try { Assert.AreEqual(1, processes.Length); Assert.AreEqual(ProcessId, processes[0].Id); Assert.IsFalse(process.HasExited); }
            finally { foreach (var current in processes) current.Dispose(); }
        }

        private static void RequireNoOutlook()
        {
            var processes = Process.GetProcessesByName("OUTLOOK");
            try { if (processes.Length != 0) Assert.Inconclusive("BLOCKED: Outlook is already running; no existing process may be used or stopped."); }
            finally { foreach (var process in processes) process.Dispose(); }
        }

        private static string HashFile(string path)
        { using (var file = File.OpenRead(path)) using (var sha = SHA256.Create()) return BitConverter.ToString(sha.ComputeHash(file)).Replace("-", ""); }
    }
}
