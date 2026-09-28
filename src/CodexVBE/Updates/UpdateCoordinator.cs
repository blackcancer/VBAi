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
        /// <summary>Synchronization object protecting timer and lifetime state.</summary>
private static readonly object timerLock = new object();
        /// <summary>Periodic timer for automatic checks in managed installations.</summary>
private static Timer timer;
        /// <summary>Cancellation source shared by automatic checks during this host lifetime.</summary>
private static CancellationTokenSource lifetime;
        /// <summary>Factory used to create a disposable release feed.</summary>
internal static Func<UpdateFeed> CreateFeed = () => new UpdateFeed();
        /// <summary>Launcher used to start the isolated installer worker.</summary>
internal static Action<bool> LaunchWorker = LaunchWorkerNative;
        /// <summary>Stores the start process used by UpdateCoordinator.</summary>
internal static Func<ProcessStartInfo, Process> StartProcess = Process.Start;
        /// <summary>Stores the create timer used by UpdateCoordinator.</summary>
internal static Func<TimerCallback, object, TimeSpan, TimeSpan, Timer> CreateTimer = NewTimer;
        /// <summary>Performs the new timer operation for UpdateCoordinator.</summary>
/// <param name="callback">The callback used by this operation.</param>
/// <param name="state">The state used by this operation.</param>
/// <param name="due">The due used by this operation.</param>
/// <param name="period">The period used by this operation.</param>
/// <returns>The result produced by this operation.</returns>
private static Timer NewTimer(TimerCallback callback, object state, TimeSpan due, TimeSpan period) => new Timer(callback, state, due, period);
        /// <summary>Serializes release checks, optionally downloads the package, and may schedule installation.</summary>
        /// <param name="automatic">Whether automatic-check preferences and skipped-version settings apply.</param>
        /// <param name="progress">Optional download progress reporter.</param><param name="ct">Cancellation token.</param>
        /// <returns>Selected release, or null when the check is not due or no eligible release exists.</returns>
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
                        bool cancelled = previous?.Completed == true && (previous.Status == "Update cancelled." || previous.Status == "Installation status is uncertain. Check the installed version.") && UpdateVersion.Parse(previous.TargetVersion).CompareTo(release.Version) == 0;
                        if (preferences.CheckAutomatically && preferences.InstallAutomatically && !cancelled && (previous == null || previous.Completed) && UpdateInstallation.IsManaged(UpdateState.InstallationDirectory))
                            Schedule(release, path, true);
                    }
                    preferences = UpdateState.Load(); preferences.LastCheckUtc = DateTime.UtcNow; UpdateState.Save(preferences);
                    return release;
                }
            }
        }
        /// <summary>Records the current managed host and starts the hourly automatic-check timer.</summary>
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
                        timer = CreateTimer(state => { _ = RunAutomatic(token); }, null, TimeSpan.Zero, TimeSpan.FromHours(1));
                    }
                }
            }
            catch (Exception) { LoadLog.Write("Update startup is unavailable."); }
        }
        /// <summary>Runs one automatic check and logs failures unless the host lifetime cancelled it.</summary>
        /// <param name="token">Host-lifetime cancellation token.</param>
        /// <returns>Task that completes when the check attempt is finished.</returns>
private static async Task RunAutomatic(CancellationToken token)
        {
            try { await Check(true, null, token); }
            catch (OperationCanceledException) when (token.IsCancellationRequested) { }
            catch (Exception) { LoadLog.Write("Automatic update check is unavailable. Use the Updates window to retry."); }
        }
        /// <summary>Stops the periodic check timer and cancels its outstanding work.</summary>
internal static void Stop()
        {
            lock (timerLock)
            {
                timer?.Dispose(); timer = null;
                lifetime?.Cancel(); lifetime?.Dispose(); lifetime = null;
            }
        }
        /// <summary>Validates a staged installer, persists a worker job, then launches the separate updater process.</summary>
        /// <param name="release">Release that supplied the installer asset.</param><param name="path">Path to the staged installer.</param>
        /// <param name="background">Whether the updater worker should run without a visible window.</param>
        /// <exception cref="InvalidOperationException">The installation is unmanaged or another job is active.</exception>
        /// <exception cref="InvalidDataException">The staged package does not match the release digest.</exception>
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
        /// <summary>Copies the signed updater beside the user cache and starts it with the selected visibility mode.</summary>
        /// <param name="background">Whether the worker should use its background mode.</param>
        /// <exception cref="FileNotFoundException">The deployed updater executable is missing.</exception>
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
            StartProcess(new ProcessStartInfo(worker, background ? "--background" : "") { UseShellExecute = true, WindowStyle = background ? ProcessWindowStyle.Hidden : ProcessWindowStyle.Normal });
        }
    }
}
