using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Web.Script.Serialization;

namespace CodexVBE
{
    /// <summary>Exécuté par le programme externe ; il ne charge pas CodexVBE.dll et ne remplace aucun fichier lui-même.</summary>
    internal sealed class UpdateInstallerRunner
    {
        internal static Func<ProcessStartInfo, Process> StartProcess = Process.Start;
        internal static Func<int, Process> ReadProcess = Process.GetProcessById;
        private readonly string root;
        internal Func<UpdateHostLease, bool> IsAlive = IsAliveNative;
        internal Func<string, bool> VerifySignature = VerifySignatureNative;
        internal Func<string, int> Install = InstallNative;
        internal Func<string, string> InstalledVersion = directory => FileVersionInfo.GetVersionInfo(Path.Combine(directory, "CodexVBE.dll")).ProductVersion;
        internal bool Installing { get; private set; }
        internal UpdateInstallerRunner(string root) { this.root = Path.GetFullPath(root); }
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
        internal bool Tick(UpdateInstallJob job)
        {
            job.Validate(root);
            if (job.Completed) return true;
            if (HostsOpen(job)) return false;
            FileStream gate;
            try { gate = new FileStream(Path.Combine(root, "installation.lock"), FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None); }
            catch (IOException) { return false; }
            using (gate)
            try
            {
                var persisted = UpdateInstallJob.Load(root);
                if (persisted != null && (persisted.Sha256 != job.Sha256 || persisted.TargetVersion != job.TargetVersion)) return true;
                if (persisted?.Completed == true) { job.Completed = true; job.Status = persisted.Status; return true; }
                // Recheck integrity and Windows trust immediately before executing cached bytes.
                if (!File.Exists(job.InstallerPath) || UpdatePaths.Hash(job.InstallerPath) != job.Sha256 || !VerifySignature(job.InstallerPath))
                    throw new InvalidDataException("Installer verification failed.");
                if (HostsOpen(job)) return false;
                Installing = true;
                job.Status = "Installing update…"; job.Save(root);
                int code = Install(job.InstallerPath);
                job.RestartRequired = code == 3010 || code == 1641;
                bool verified = job.RestartRequired || (code == 0 && UpdateVersion.Parse(InstalledVersion(job.InstallationDirectory))?.CompareTo(UpdateVersion.Parse(job.TargetVersion)) == 0);
                job.Status = job.RestartRequired ? "Update installed. Restart Windows to finish." : verified ? "Update installed. Restart the VBA host." : "Installation failed. Check the installer log.";
                job.Completed = true; job.Save(root);
            }
            catch (Exception)
            {
                job.Status = "Installation failed. Check the installer log."; job.Completed = true; job.Save(root);
            }
            finally { Installing = false; }
            return true;
        }
        private static bool IsAliveNative(UpdateHostLease lease)
        {
            try { using (var process = ReadProcess(lease.Pid)) return !process.HasExited && process.StartTime.ToUniversalTime().Ticks == lease.StartTimeUtcTicks; }
            catch (ArgumentException) { return false; }
            catch (InvalidOperationException) { return false; }
            // Access denied is deliberately propagated to HostsOpen, which waits.
        }
        private static int InstallNative(string path)
        {
            bool msi = path.EndsWith(".msi", StringComparison.OrdinalIgnoreCase);
            var start = new ProcessStartInfo(msi ? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System), "msiexec.exe") : path,
                msi ? "/i \"" + path + "\" /quiet /norestart /L*v \"" + path + ".install.log\"" : "/update /quiet /norestart") { UseShellExecute = true, WindowStyle = ProcessWindowStyle.Hidden };
            using (var process = StartProcess(start))
            {
                if (process == null) throw new InvalidOperationException("Installer did not start.");
                process.WaitForExit(); return process.ExitCode;
            }
        }
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
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct TrustFile { internal uint Size; [MarshalAs(UnmanagedType.LPWStr)] internal string Path; internal IntPtr File, Subject; }
        [StructLayout(LayoutKind.Sequential)]
        private struct TrustData
        {
            internal uint Size; internal IntPtr Policy, Sip; internal uint UiChoice, RevocationChecks, UnionChoice;
            internal IntPtr File; internal uint StateAction; internal IntPtr StateData, Url; internal uint Flags, Context;
        }
        [DllImport("wintrust.dll", ExactSpelling = true, CharSet = CharSet.Unicode)]
        private static extern int WinVerifyTrust(IntPtr window, ref Guid action, ref TrustData data);
    }
}
