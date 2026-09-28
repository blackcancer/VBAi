using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Web.Script.Serialization;

namespace CodexVBE
{
    /// <summary>Persisted choices and timestamps that control automatic release checks and installation.</summary>
    internal sealed class UpdatePreferences
    {
        /// <summary>Gets or sets whether periodic release checks are enabled.</summary><value>True by default.</value>
        public bool CheckAutomatically { get; set; } = true;
        /// <summary>Gets or sets whether eligible updates should be downloaded without prompting.</summary><value>True by default.</value>
        public bool DownloadAutomatically { get; set; } = true;
        /// <summary>Gets or sets whether downloaded updates should be installed automatically.</summary><value>False by default.</value>
        public bool InstallAutomatically { get; set; }
        /// <summary>Gets or sets whether prerelease versions are eligible.</summary><value>False by default.</value>
        public bool IncludePrereleases { get; set; }
        /// <summary>Gets or sets a version the user chose to skip.</summary><value>Skipped version text, or null.</value>
        public string SkippedVersion { get; set; }
        /// <summary>Gets or sets the UTC time of the latest release check.</summary><value>Last check timestamp in UTC.</value>
        public DateTime LastCheckUtc { get; set; }
        /// <summary>Gets or sets the UTC time of the latest download or install attempt.</summary><value>Last attempt timestamp in UTC.</value>
        public DateTime LastAttemptUtc { get; set; }
    }
    /// <summary>Contrat écrit par l’installeur dans son répertoire de déploiement.</summary>
    internal sealed class UpdateInstallation
    {
        /// <summary>Gets or sets the product marker used to verify installation ownership.</summary><value>Expected product identifier.</value>
        public string Product { get; set; }
        /// <summary>Gets or sets the unique installation identifier.</summary><value>GUID text written by the installer.</value>
        public string InstallationId { get; set; }
        /// <summary>Gets or sets the installed process architecture.</summary><value>Installer architecture marker.</value>
        public string Architecture { get; set; }
        /// <summary>Gets or sets the updater protocol version implemented by the installation.</summary><value>Protocol number.</value>
        public int UpdateProtocol { get; set; }
        /// <summary>Checks for a recognized VBAi x64 installation marker using update protocol 1.</summary>
        /// <param name="directory">Installation directory containing <c>vbai-installation.json</c>.</param><returns>Whether the marker passes every required check.</returns>
        internal static bool IsManaged(string directory)
        {
            try
            {
                var marker = new JavaScriptSerializer().Deserialize<UpdateInstallation>(File.ReadAllText(Path.Combine(directory, "vbai-installation.json")));
                return marker != null && marker.Product == "VBAi" && marker.Architecture == "win-x64" && marker.UpdateProtocol == 1 && Guid.TryParse(marker.InstallationId, out var id);
            }
            catch (Exception) { return false; }
        }
    }
    /// <summary>Loads updater preferences, caches release metadata, and records installation host leases.</summary>
    internal static class UpdateState
    {
        /// <summary>Optional test override for the add-in installation directory.</summary>
        internal static string InstallationDirectoryOverride;
        /// <summary>Gets the configured override or the directory containing the add-in assembly.</summary><value>Full installation directory.</value>
        internal static string InstallationDirectory => InstallationDirectoryOverride ?? Path.GetDirectoryName(typeof(UpdateState).Assembly.Location);
        /// <summary>Stores the read product assembly used by UpdateState.</summary>
        internal static Func<Assembly> ReadProductAssembly = () => typeof(UpdateState).Assembly;
        /// <summary>Gets the assembly informational version, falling back to its assembly version.</summary><value>Current product version string.</value>
        internal static string ProductVersion { get { var assembly = ReadProductAssembly(); return assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion ?? assembly.GetName().Version.ToString(); } }
        /// <summary>Loads persisted updater preferences, returning defaults if the file is missing or invalid.</summary>
        /// <param name="root">Preference directory, or null to use <see cref="UpdatePaths.Root"/>.</param><returns>Loaded preferences or defaults.</returns>
        internal static UpdatePreferences Load(string root = null)
        {
            try { return new JavaScriptSerializer().Deserialize<UpdatePreferences>(File.ReadAllText(Path.Combine(root ?? UpdatePaths.Root, "preferences.json"))) ?? new UpdatePreferences(); }
            catch (Exception) { return new UpdatePreferences(); }
        }
        /// <summary>Serializes updater preferences to the selected root using an atomic write.</summary>
        /// <param name="preferences">Preferences to save.</param><param name="root">Preference directory, or null to use the default root.</param>
        internal static void Save(UpdatePreferences preferences, string root = null) => UpdatePaths.WriteAtomic(Path.Combine(root ?? UpdatePaths.Root, "preferences.json"), new JavaScriptSerializer().Serialize(preferences));
        /// <summary>Stores release metadata for offline display, limiting notes length and excluding credentials or signed URLs.</summary>
        /// <param name="release">Release metadata to cache.</param>
        internal static void CacheRelease(UpdateRelease release)
        {
            // Keep release notes available even if a later download fails; never cache an access token or signed CDN URL.
            if (release?.body?.Length > 64000)
                release = new UpdateRelease { tag_name = release.tag_name, body = release.body.Substring(0, 64000),
                    draft = release.draft, prerelease = release.prerelease, assets = release.assets };
            UpdatePaths.WriteAtomic(Path.Combine(UpdatePaths.Root, "latest-release.json"), new JavaScriptSerializer().Serialize(release));
        }
        /// <summary>Loads a cached release when it is newer than the current product and matches prerelease preferences.</summary>
        /// <returns>Eligible cached release, or null when unavailable or ineligible.</returns>
        internal static UpdateRelease CachedRelease()
        {
            try
            {
                string path = Path.Combine(UpdatePaths.Root, "latest-release.json");
                if (new FileInfo(path).Length > 262144) return null;
                var release = new JavaScriptSerializer().Deserialize<UpdateRelease>(File.ReadAllText(path));
                var preferences = Load();
                return release?.Version != null && !release.draft && release.Version.CompareTo(UpdateVersion.Parse(ProductVersion)) > 0 &&
                    (preferences.IncludePrereleases || (!release.prerelease && !release.Version.IsPreview)) ? release : null;
            }
            catch (Exception) { return null; }
        }
        /// <summary>Checks the 24-hour release-check and one-hour attempt intervals against the supplied UTC time.</summary>
        /// <param name="preferences">Current updater preferences and timestamps.</param><param name="now">Current time used for the comparison.</param>
        /// <returns>Whether an automatic check is due.</returns>
        internal static bool Due(UpdatePreferences preferences, DateTime now) => preferences.CheckAutomatically &&
            (preferences.LastCheckUtc > now || now - preferences.LastCheckUtc >= TimeSpan.FromHours(24)) &&
            (preferences.LastAttemptUtc > now || now - preferences.LastAttemptUtc >= TimeSpan.FromHours(1));
        /// <summary>Writes a process lease containing its PID, start time, and installation directory.</summary>
        /// <param name="root">Lease root, or null to use the default update directory.</param>
        internal static void RecordHost(string root = null)
        {
            using (var process = Process.GetCurrentProcess())
            {
                var lease = new UpdateHostLease { Pid = process.Id, StartTimeUtcTicks = process.StartTime.ToUniversalTime().Ticks, InstallationDirectory = InstallationDirectory };
                UpdatePaths.WriteAtomic(Path.Combine(root ?? UpdatePaths.Root, "hosts", process.Id + ".json"), new JavaScriptSerializer().Serialize(lease));
            }
            // Keep the lease after COM disconnection: the CLR can retain the DLL until process exit.
        }
    }
}
