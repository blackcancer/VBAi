using System;
using System.IO;
using System.Web.Script.Serialization;

namespace CodexVBE
{
    internal sealed class UpdateHostLease
    {
        public int Pid { get; set; }
        public long StartTimeUtcTicks { get; set; }
        public string InstallationDirectory { get; set; }
    }
    internal sealed class UpdateInstallJob
    {
        public string InstallerPath { get; set; }
        public string Sha256 { get; set; }
        public string TargetVersion { get; set; }
        public string InstallationDirectory { get; set; }
        public string Culture { get; set; }
        public string Status { get; set; } = "Waiting for VBA hosts to close.";
        public bool Completed { get; set; }
        public bool RestartRequired { get; set; }
        internal void Validate(string root)
        {
            if (!UpdatePaths.IsHash(Sha256) || UpdateVersion.Parse(TargetVersion) == null ||
                string.IsNullOrWhiteSpace(InstallationDirectory) || !Path.IsPathRooted(InstallationDirectory) ||
                string.IsNullOrWhiteSpace(InstallerPath) ||
                !UpdatePaths.SamePath(InstallerPath, UpdatePaths.AssetPath(root, Sha256, Path.GetFileName(InstallerPath))))
                throw new InvalidDataException("Invalid update job.");
        }
        internal void Save(string root) => UpdatePaths.WriteAtomic(Path.Combine(root, "pending.json"), new JavaScriptSerializer().Serialize(this));
        internal static UpdateInstallJob Load(string root)
        {
            string path = Path.Combine(root, "pending.json");
            if (!File.Exists(path)) return null;
            if (new FileInfo(path).Length > 16384) throw new InvalidDataException("Invalid update job size.");
            var job = new JavaScriptSerializer().Deserialize<UpdateInstallJob>(File.ReadAllText(path));
            if (job == null) throw new InvalidDataException("Missing update job.");
            job.Validate(root); return job;
        }
    }
}
