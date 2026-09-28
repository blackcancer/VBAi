using System;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace CodexVBE
{
    /// <summary>Coordonne les vérifications entre les hôtes et l’application différée par un programme distinct.</summary>
    internal static class UpdateCoordinator
    {
        private static readonly object timerLock = new object();
        private static Timer timer;
        private static CancellationTokenSource lifetime;
        internal static Func<UpdateFeed> CreateFeed = () => new UpdateFeed();
        internal static Action<bool> LaunchWorker = LaunchWorkerNative;
        internal static async Task<UpdateRelease> Check(bool automatic, IProgress<int> progress, CancellationToken ct)
        {
            string root = UpdatePaths.Root;
            Directory.CreateDirectory(root);
            FileStream gate;
            try { gate = new FileStream(Path.Combine(root, "operation.lock"), FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None); }
            catch (IOException) { throw new InvalidOperationException("Another VBA host is checking for updates."); }
            using (gate)
            {
                var preferences = UpdateState.Load();
                if (automatic && !UpdateState.Due(preferences, DateTime.UtcNow)) return null;
                preferences.LastAttemptUtc = DateTime.UtcNow; UpdateState.Save(preferences);
                using (var feed = CreateFeed())
                {
                    var release = await feed.Check(UpdateVersion.Parse(UpdateState.ProductVersion), preferences.IncludePrereleases,
                        automatic ? preferences.SkippedVersion : null, ct);
                    UpdateState.CacheRelease(release);
                    if (automatic && release?.Installer != null && preferences.DownloadAutomatically)
                    {
                        string path = await feed.Download(release.Installer, root, progress, ct);
                        // Re-read user preferences after a potentially long download.
                        preferences = UpdateState.Load();
                        var previous = UpdateInstallJob.Load(root);
                        bool cancelled = previous?.Completed == true && (previous.Status == "Update cancelled." || previous.Status == "Installation status is uncertain. Check the installed version.") && UpdateVersion.Parse(previous.TargetVersion)?.CompareTo(release.Version) == 0;
                        if (preferences.CheckAutomatically && preferences.InstallAutomatically && !cancelled && (previous == null || previous.Completed) && UpdateInstallation.IsManaged(UpdateState.InstallationDirectory))
                            Schedule(release, path, true);
                    }
                    preferences = UpdateState.Load(); preferences.LastCheckUtc = DateTime.UtcNow; UpdateState.Save(preferences);
                    return release;
                }
            }
        }
        internal static void Start()
        {
            // Development checkouts never start network traffic or an installer automatically.
            if (!UpdateInstallation.IsManaged(UpdateState.InstallationDirectory)) return;
            try
            {
                UpdateState.RecordHost();
                var pending = UpdateInstallJob.Load(UpdatePaths.Root);
                if (pending != null && !pending.Completed) LaunchWorker(true);
                lock (timerLock)
                {
                    if (timer == null)
                    {
                        lifetime = new CancellationTokenSource(); var token = lifetime.Token;
                        timer = new Timer(state => { _ = RunAutomatic(token); }, null, TimeSpan.Zero, TimeSpan.FromHours(1));
                    }
                }
            }
            catch (Exception) { LoadLog.Write("Update startup is unavailable."); }
        }
        private static async Task RunAutomatic(CancellationToken token)
        {
            try { await Check(true, null, token); }
            catch (OperationCanceledException) when (token.IsCancellationRequested) { }
            catch (Exception) { LoadLog.Write("Automatic update check is unavailable. Use the Updates window to retry."); }
        }
        internal static void Stop()
        {
            lock (timerLock)
            {
                timer?.Dispose(); timer = null;
                lifetime?.Cancel(); lifetime?.Dispose(); lifetime = null;
            }
        }
        internal static void Schedule(UpdateRelease release, string path, bool background)
        {
            if (!UpdateInstallation.IsManaged(UpdateState.InstallationDirectory)) throw new InvalidOperationException("An installer-managed deployment is required.");
            if (release?.Installer?.Hash == null || !UpdatePaths.SamePath(path, UpdatePaths.AssetPath(UpdatePaths.Root, release.Installer.Hash, release.Installer.name)))
                throw new InvalidDataException("Invalid staged update.");
            using (var gate = new FileStream(Path.Combine(UpdatePaths.Root, "installation.lock"), FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None))
            {
                if (!File.Exists(path) || UpdatePaths.Hash(path) != release.Installer.Hash) throw new InvalidDataException("The staged installer has changed.");
                var existing = UpdateInstallJob.Load(UpdatePaths.Root);
                if (existing != null && !existing.Completed) throw new InvalidOperationException("An update is already scheduled.");
                UpdateState.RecordHost();
                var job = new UpdateInstallJob { InstallerPath = path, Sha256 = release.Installer.Hash, TargetVersion = release.Version.Text,
                    InstallationDirectory = UpdateState.InstallationDirectory, Culture = UiText.Culture.Name };
                job.Validate(UpdatePaths.Root); job.Save(UpdatePaths.Root);
            }
            LaunchWorker(background);
        }
        private static void LaunchWorkerNative(bool background)
        {
            string source = Path.Combine(UpdateState.InstallationDirectory, "VBAi.Updater.exe");
            if (!File.Exists(source)) throw new FileNotFoundException("The updater is missing.");
            string directory = Path.Combine(UpdatePaths.Root, "tools", UpdatePaths.Hash(source));
            Directory.CreateDirectory(directory);
            string worker = Path.Combine(directory, "VBAi.Updater.exe");
            if (!File.Exists(worker)) File.Copy(source, worker);
            else if (UpdatePaths.Hash(worker) != UpdatePaths.Hash(source)) throw new InvalidDataException("Invalid cached updater.");
            string config = source + ".config"; if (File.Exists(config) && !File.Exists(worker + ".config")) File.Copy(config, worker + ".config");
            Process.Start(new ProcessStartInfo(worker, background ? "--background" : "") { UseShellExecute = true, WindowStyle = background ? ProcessWindowStyle.Hidden : ProcessWindowStyle.Normal });
        }
    }
}
