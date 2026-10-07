using System.IO;
using System.Web.Script.Serialization;

namespace VBAi
{

    /// <summary>Records a VBA host process that must close before the installation can be replaced.</summary>
    internal sealed class UpdateHostLease
    {

        /// <summary>Gets or sets the host process identifier.</summary><value>Process ID that currently owns the installation.</value>
        public int Pid { get; set; }

        /// <summary>Gets or sets the host process start time as UTC ticks.</summary><value>Start timestamp used to detect PID reuse.</value>
        public long StartTimeUtcTicks { get; set; }

        /// <summary>Gets or sets the installation directory loaded by the host.</summary><value>Full path to the installed add-in.</value>
        public string InstallationDirectory { get; set; }
    }

    /// <summary>Persisted request for the updater worker to install a staged package after hosts exit.</summary>
    internal sealed class UpdateInstallJob
    {

        /// <summary>Gets or sets the staged installer path.</summary><value>Path inside the hash-named update cache.</value>
        public string InstallerPath { get; set; }

        /// <summary>Gets or sets the SHA-256 digest expected for the staged package.</summary><value>Lowercase hexadecimal digest.</value>
        public string Sha256 { get; set; }

        /// <summary>Gets or sets the product version requested by the job.</summary><value>Target semantic version text.</value>
        public string TargetVersion { get; set; }

        /// <summary>Gets or sets the installation directory to update.</summary><value>Full path to the managed deployment.</value>
        public string InstallationDirectory { get; set; }

        /// <summary>Gets or sets the culture used for worker messages.</summary><value>Culture name.</value>
        public string Culture { get; set; }

        /// <summary>Gets or sets user-facing worker progress text.</summary><value>Current status message.</value>
        public string Status { get; set; } = "Waiting for VBA hosts to close.";

        /// <summary>Gets or sets whether the worker has reached a final state.</summary><value>Completion flag.</value>
        public bool Completed { get; set; }

        /// <summary>Whether the completed installer result was verified as successful.</summary>
        /// <value>True only when the completed installer exit/version or reboot-required outcome was accepted; false for failure, cancellation or an uncertain attempt.</value>
        public bool Succeeded { get; set; }

        /// <summary>Gets or sets whether the installer reported a required Windows restart.</summary><value>True for installer exit codes 3010 or 1641; a normal success separately asks the user to restart the VBA host.</value>
        public bool RestartRequired { get; set; }

        /// <summary>Validates the target package, version, installation path, and cache containment.</summary>
        /// <param name="root">Update cache root against which the installer path is checked.</param>
        /// <exception cref="InvalidDataException">A required field or asset path is invalid.</exception>
        internal void Validate(string root)
        {
            if (!UpdatePaths.IsHash(Sha256) || UpdateVersion.Parse(TargetVersion) == null ||
                string.IsNullOrWhiteSpace(InstallationDirectory) || !Path.IsPathRooted(InstallationDirectory) ||
                string.IsNullOrWhiteSpace(InstallerPath) ||
                !UpdatePaths.SamePath(InstallerPath, UpdatePaths.AssetPath(root, Sha256, Path.GetFileName(InstallerPath))))
                throw new InvalidDataException("Invalid update job.");
        }

        /// <summary>Atomically writes this job to the worker's pending-job file.</summary><param name="root">Update cache root.</param>
        internal void Save(string root) => UpdatePaths.WriteAtomic(Path.Combine(root, "pending.json"), new JavaScriptSerializer().Serialize(this));

        /// <summary>Loads and validates the pending worker job, returning null when no job exists.</summary>
        /// <param name="root">Update cache root.</param><returns>Validated job, or null when absent.</returns>
        /// <exception cref="InvalidDataException">The file is oversized or contains an invalid job.</exception>
        internal static UpdateInstallJob Load(string root)
        {
            string path = Path.Combine(root, "pending.json");
            if (!File.Exists(path)) return null;
            if (new FileInfo(path).Length > 16384) throw new InvalidDataException("Invalid update job size.");
            var job = new JavaScriptSerializer().Deserialize<UpdateInstallJob>(File.ReadAllText(path)) ?? throw new InvalidDataException("Missing update job.");
            job.Validate(root); return job;
        }
    }
}
