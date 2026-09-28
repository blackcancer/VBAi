using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Web.Script.Serialization;

namespace CodexVBE
{
    internal sealed class UpdatePreferences
    {
        public bool CheckAutomatically { get; set; } = true;
        public bool DownloadAutomatically { get; set; } = true;
        public bool InstallAutomatically { get; set; }
        public bool IncludePrereleases { get; set; }
        public string SkippedVersion { get; set; }
        public DateTime LastCheckUtc { get; set; }
        public DateTime LastAttemptUtc { get; set; }
    }
    /// <summary>Contrat écrit par l’installeur dans son répertoire de déploiement.</summary>
    internal sealed class UpdateInstallation
    {
        public string Product { get; set; }
        public string InstallationId { get; set; }
        public string Architecture { get; set; }
        public int UpdateProtocol { get; set; }
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
    internal static class UpdateState
    {
        internal static string InstallationDirectoryOverride;
        internal static string InstallationDirectory => InstallationDirectoryOverride ?? Path.GetDirectoryName(typeof(UpdateState).Assembly.Location);
        internal static string ProductVersion => typeof(UpdateState).Assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion ?? typeof(UpdateState).Assembly.GetName().Version.ToString();
        internal static UpdatePreferences Load(string root = null)
        {
            try { return new JavaScriptSerializer().Deserialize<UpdatePreferences>(File.ReadAllText(Path.Combine(root ?? UpdatePaths.Root, "preferences.json"))) ?? new UpdatePreferences(); }
            catch (Exception) { return new UpdatePreferences(); }
        }
        internal static void Save(UpdatePreferences preferences, string root = null) => UpdatePaths.WriteAtomic(Path.Combine(root ?? UpdatePaths.Root, "preferences.json"), new JavaScriptSerializer().Serialize(preferences));
        internal static void CacheRelease(UpdateRelease release)
        {
            // Keep release notes available even if a later download fails; never cache an access token or signed CDN URL.
            if (release?.body?.Length > 64000)
                release = new UpdateRelease { tag_name = release.tag_name, body = release.body.Substring(0, 64000),
                    draft = release.draft, prerelease = release.prerelease, assets = release.assets };
            UpdatePaths.WriteAtomic(Path.Combine(UpdatePaths.Root, "latest-release.json"), new JavaScriptSerializer().Serialize(release));
        }
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
        internal static bool Due(UpdatePreferences preferences, DateTime now) => preferences.CheckAutomatically &&
            (preferences.LastCheckUtc > now || now - preferences.LastCheckUtc >= TimeSpan.FromHours(24)) &&
            (preferences.LastAttemptUtc > now || now - preferences.LastAttemptUtc >= TimeSpan.FromHours(1));
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
