using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Web.Script.Serialization;

namespace VBAi
{

    /// <summary>Exécuté par le programme externe ; il ne charge pas VBAi.dll et ne remplace aucun fichier lui-même.</summary>
    internal sealed class UpdateInstallerRunner
    {

        /// <summary>Maintains the start process state for update installer runner.</summary>
        internal static Func<ProcessStartInfo, Process> StartProcess = Process.Start;

        /// <summary>Maintains the read process state for update installer runner.</summary>
        internal static Func<int, Process> ReadProcess = Process.GetProcessById;

        /// <summary>Update root containing the pending job, staged package, and host leases.</summary>
        private readonly string root;

        /// <summary>Process-liveness check used to detect whether a registered host remains active.</summary>
        internal Func<UpdateHostLease, bool> IsAlive = IsAliveNative;

        /// <summary>Authenticode verification delegate for the staged installer.</summary>
        internal Func<string, bool> VerifySignature = VerifySignatureNative;

        /// <summary>Product publisher policy is separate from general Windows signature trust.</summary>
        internal Func<string, bool> VerifyPublisher = UpdatePublisherPolicy.Accepts;

        /// <summary>Maintains the wait for installer state for update installer runner.</summary>
        internal static Func<Process, int, bool> WaitForInstaller = (process, milliseconds) => process.WaitForExit(milliseconds);

        /// <summary>Maintains the installer timeout milliseconds state for update installer runner.</summary>
        internal const int InstallerTimeoutMilliseconds = 15 * 60 * 1000;

        /// <summary>Maintains the uncertain status state for update installer runner.</summary>
        internal const string UncertainStatus = "Installation outcome is uncertain. Check the installer and its log before attempting another update.";

        /// <summary>Delegate that launches an installer and returns its process exit code.</summary>
        internal Func<string, int> Install = InstallNative;

        /// <summary>Delegate that reads the installed product version from a directory.</summary>
        internal Func<string, string> InstalledVersion = directory => FileVersionInfo.GetVersionInfo(Path.Combine(directory, "VBAi.dll")).ProductVersion;

        /// <summary>Gets whether the installer process is currently running.</summary><value>True while Install is executing.</value>
        internal bool Installing { get; private set; }

        /// <summary>Creates a runner bound to one normalized update root.</summary><param name="root">Update cache root path.</param>
        internal UpdateInstallerRunner(string root) { this.root = Path.GetFullPath(root); }

        /// <summary>Checks host leases and treats malformed leases as potentially active hosts.</summary>
        /// <param name="job">Job whose installation directory must be free of loaded hosts.</param><returns>Whether a matching or unverifiable host may still be open.</returns>
        internal bool HostsOpen(UpdateInstallJob job)
        {
            string hosts = Path.Combine(root, "hosts");
            if (!Directory.Exists(hosts)) return false;
            foreach (string file in Directory.GetFiles(hosts, "*.json"))
            {
                try
                {
                    var lease = new JavaScriptSerializer().Deserialize<UpdateHostLease>(File.ReadAllText(file));
                    if (lease == null || lease.Pid <= 0 || lease.StartTimeUtcTicks <= 0 || string.IsNullOrEmpty(lease.InstallationDirectory)) throw new InvalidDataException();
                    if (IsAlive(lease))
                    {
                        if (UpdatePaths.SamePath(lease.InstallationDirectory, job.InstallationDirectory)) return true;
                    }
                    else File.Delete(file);
                }
                catch (Exception) { return true; } // Cannot prove a potentially loaded host is closed.
            }
            return false;
        }

        /// <summary>Advances the job after host closure and revalidates package integrity and trust before installation.</summary>
        /// <param name="job">Job to validate and process.</param><returns>True when complete or superseded; false while a host or another worker remains active.</returns>
        internal bool Tick(UpdateInstallJob job)
        {
            job.Validate(root);
            if (job.Completed) return true;
            FileStream gate;
            try { gate = new FileStream(Path.Combine(root, "installation.lock"), FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None); }
            catch (IOException) { return false; }
            using (gate)
            try
            {
                var persisted = UpdateInstallJob.Load(root);
                if (persisted != null && (persisted.Sha256 != job.Sha256 || persisted.TargetVersion != job.TargetVersion)) return true;
                if (persisted?.Completed == true) { job.Completed = true; job.Succeeded = persisted.Succeeded; job.Status = persisted.Status; return true; }
                job.Succeeded = false;
                string attempt = Path.Combine(root, "installation.started");
                if (File.Exists(attempt))
                {
                    job.Status = UncertainStatus; job.Completed = true; job.Save(root); return true;
                }
                if (!VerifyPublisher(job.InstallerPath))
                {
                    job.Status = "Installation refused: no trusted VBAi publisher is configured or the publisher does not match.";
                    job.Completed = true; job.Save(root); return true;
                }
                if (HostsOpen(job)) return false;
                // Recheck integrity and Windows trust immediately before executing cached bytes.
                if (!File.Exists(job.InstallerPath) || UpdatePaths.Hash(job.InstallerPath) != job.Sha256 || !VerifySignature(job.InstallerPath))
                    throw new InvalidDataException("Installer verification failed.");
                if (HostsOpen(job)) return false;
                Installing = true;
                job.Status = "Installing update…"; job.Save(root);
                // Persist before process creation. A crash or timeout must never cause an automatic replay.
                using (var marker = new FileStream(attempt, FileMode.CreateNew, FileAccess.Write, FileShare.Read))
                {
                    byte[] identity = System.Text.Encoding.UTF8.GetBytes(DateTime.UtcNow.ToString("O") + "\n" + job.Sha256 + "\n" + job.TargetVersion);
                    marker.Write(identity, 0, identity.Length); marker.Flush(true);
                }
                int code = Install(job.InstallerPath);
                job.RestartRequired = code == 3010 || code == 1641;
                bool verified = job.RestartRequired || (code == 0 && UpdateVersion.Parse(InstalledVersion(job.InstallationDirectory))?.CompareTo(UpdateVersion.Parse(job.TargetVersion)) == 0);
                job.Succeeded = verified;
                job.Status = job.RestartRequired ? "Update installed. Restart Windows to finish." : verified ? "Update installed. Restart the VBA host." : "Installation failed. Check the installer log.";
                job.Completed = true; job.Save(root);
                File.Delete(attempt);
            }
            catch (Exception)
            {
                job.Succeeded = false;
                job.Status = File.Exists(Path.Combine(root, "installation.started")) ? UncertainStatus : "Installation failed. Check the installer log.";
                job.Completed = true; job.Save(root);
            }
            finally { Installing = false; }
            return true;
        }

        /// <summary>Checks process identity using PID and UTC start time to avoid PID reuse.</summary>
        /// <param name="lease">Persisted host lease.</param><returns>Whether the same process instance is alive.</returns>
        private static bool IsAliveNative(UpdateHostLease lease)
        {
            try { using (var process = ReadProcess(lease.Pid)) return !process.HasExited && process.StartTime.ToUniversalTime().Ticks == lease.StartTimeUtcTicks; }
            catch (ArgumentException) { return false; }
            catch (InvalidOperationException) { return false; }
            // Access denied is deliberately propagated to HostsOpen, which waits.
        }

        /// <summary>Runs an MSI through msiexec or starts the setup executable with quiet update flags.</summary>
        /// <param name="path">Verified installer path.</param><returns>Installer process exit code.</returns>
        /// <exception cref="InvalidOperationException">The installer process could not start.</exception>
        private static int InstallNative(string path)
        {
            bool msi = path.EndsWith(".msi", StringComparison.OrdinalIgnoreCase);
            var start = new ProcessStartInfo(msi ? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System), "msiexec.exe") : path,
                msi ? "/i \"" + path + "\" /quiet /norestart /L*v \"" + path + ".install.log\"" : "/update /quiet /norestart") { UseShellExecute = true, WindowStyle = ProcessWindowStyle.Hidden };
            using (var process = StartProcess(start))
            {
                if (process == null) throw new InvalidOperationException("Installer did not start.");
                if (!WaitForInstaller(process, InstallerTimeoutMilliseconds))
                    throw new TimeoutException(UncertainStatus);
                return process.ExitCode;
            }
        }

        /// <summary>Uses WinVerifyTrust to verify the installer without UI and then releases trust state.</summary>
        /// <param name="path">Installer file to verify.</param><returns>Whether WinTrust accepts the file.</returns>
        internal static bool VerifySignatureNative(string path)
        {
            var file = new TrustFile { Size = (uint)Marshal.SizeOf(typeof(TrustFile)), Path = path };
            IntPtr pointer = Marshal.AllocHGlobal(Marshal.SizeOf(typeof(TrustFile)));
            Marshal.StructureToPtr(file, pointer, false);
            var data = new TrustData { Size = (uint)Marshal.SizeOf(typeof(TrustData)), UiChoice = 2, UnionChoice = 1, File = pointer, StateAction = 1 };
            var action = new Guid("00AAC56B-CD44-11D0-8CC2-00C04FC295EE");
            try { return WinVerifyTrust(new IntPtr(-1), ref action, ref data) == 0; }
            finally { data.StateAction = 2; WinVerifyTrust(new IntPtr(-1), ref action, ref data); Marshal.DestroyStructure(pointer, typeof(TrustFile)); Marshal.FreeHGlobal(pointer); }
        }

        /// <summary>WINTRUST_FILE_INFO-compatible data identifying the file to verify.</summary>
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct TrustFile {

/// <summary>Structure size, file path, optional file handle, and subject pointer.</summary>
internal uint Size;

/// <summary>Unicode path of the installer file being verified.</summary>
[MarshalAs(UnmanagedType.LPWStr)] internal string Path;

/// <summary>Optional open file handle and subject pointer, unused by this caller.</summary>
internal IntPtr File, Subject; }

        /// <summary>WINTRUST_DATA-compatible policy, file choice, and trust-state data.</summary>
        [StructLayout(LayoutKind.Sequential)]
        private struct TrustData
        {

            /// <summary>Structure size in bytes.</summary>
            internal uint Size;

/// <summary>Optional policy and SIP provider data pointers.</summary>
internal IntPtr Policy, Sip;

/// <summary>UI mode, revocation-check setting, and selected union member.</summary>
internal uint UiChoice, RevocationChecks, UnionChoice;

            /// <summary>Pointer to the selected trust file description.</summary>
            internal IntPtr File;

/// <summary>Trust state action used to open or close provider state.</summary>
internal uint StateAction;

/// <summary>Provider state, optional URL, flags, and caller context.</summary>
internal IntPtr StateData, Url;

/// <summary>Trust-provider behavior flags and caller context pointer.</summary>
internal uint Flags, Context;
        }

        /// <summary>Verifies a file's Authenticode trust state using the Windows trust provider.</summary>
        /// <param name="window">Window handle for trust-provider UI, unused with no-UI policy.</param><param name="action">Trust action GUID.</param>
        /// <param name="data">Trust policy and file data.</param><returns>Zero when the file is trusted.</returns>
        [DllImport("wintrust.dll", ExactSpelling = true, CharSet = CharSet.Unicode)]
        private static extern int WinVerifyTrust(IntPtr window, ref Guid action, ref TrustData data);
    }
}
