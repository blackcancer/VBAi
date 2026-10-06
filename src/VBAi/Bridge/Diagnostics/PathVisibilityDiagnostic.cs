using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Web.Script.Serialization;

namespace VBAi
{

    /// <summary>Explicit host-environment opt-in with a fixed synthetic allowlist and exact owner-STA guard.</summary>
    internal sealed class PathVisibilityDiagnostic
    {

        /// <summary>Opt-in environment variable containing the bounded synthetic-path manifest file path.</summary>
        internal const string EnvironmentName = "VBAi_TEST_PATH_VISIBILITY_MANIFEST";

        /// <summary>Exact bridge command that opts into the synthetic path observation.</summary>
        internal const string CommandName = "diagnostic_path_visibility";

        /// <summary>Fixed filename stem used below each GUID-owned synthetic directory.</summary>
        internal const string SyntheticName = "VBAi.PathVisibility.synthetic";

        /// <summary>Connected host process ID captured when the diagnostic object is created.</summary>
        private readonly int ownerPid;

        /// <summary>Native VBE UI thread ID captured at diagnostic construction.</summary>
        private readonly uint ownerTid;

        /// <summary>Validated two-directory/two-file synthetic allowlist, or null when the diagnostic was not enabled.</summary>
        private readonly string[] paths;

        /// <summary>Safe type-and-message summary retained when the opt-in manifest fails validation.</summary>
        private readonly string preparationError;

        /// <summary>Captures the owner IDs and validates the opt-in manifest without reading any listed path contents.</summary>
        /// <param name="processId">Connected host process ID expected for later owner checks.</param>
        internal PathVisibilityDiagnostic(int processId)
        {
            ownerPid = processId; ownerTid = PathVisibilityObservation.ThreadId;
            string manifest = Environment.GetEnvironmentVariable(EnvironmentName);
            if (string.IsNullOrWhiteSpace(manifest)) return;
            try
            {
                ExcelLocalPath(manifest);
                string name = Path.GetFileName(manifest);
                const string suffix = ".visibility.json";
                Guid guid;
                if (!name.EndsWith(suffix, StringComparison.Ordinal) ||
                    !Guid.TryParseExact(name.Substring(0, name.Length - suffix.Length), "N", out guid))
                    throw new ArgumentException("A GUID-owned manifest filename is required.");
                var drive = new DriveInfo(Path.GetPathRoot(manifest));
                if (drive.DriveType != DriveType.Fixed && drive.DriveType != DriveType.Removable && drive.DriveType != DriveType.Ram)
                    throw new ArgumentException("The diagnostic manifest must be on a local drive.");
                RejectReparseAncestors(manifest);
                using (var input = new FileStream(manifest, FileMode.Open, FileAccess.Read, FileShare.Read))
                {
                    if (input.Length < 2 || input.Length > 4096) throw new ArgumentException("Manifest size is invalid.");
                    using (var reader = new StreamReader(input, new System.Text.UTF8Encoding(false, true), false))
                        paths = ValidateManifest(reader.ReadToEnd(), Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), Path.GetTempPath());
                }
            }
            catch (Exception error) { preparationError = error.GetType().Name + ": " + error.Message; }
        }

        /// <summary>Runs only on the captured process's native STA and returns attribute/token observations after manifest validation.</summary>
        /// <returns>Diagnostic evidence tagged with the production assembly MVID; disabled or invalid manifests throw.</returns>
        internal IDictionary<string, object> Read()
        {
            RequireOwner(ownerPid, ownerTid, PathVisibilityObservation.ProcessId, PathVisibilityObservation.ThreadId, Thread.CurrentThread.GetApartmentState());
            if (paths == null) throw new InvalidOperationException(preparationError == null ? "Path visibility diagnostic is disabled." : "Invalid diagnostic manifest: " + preparationError);
            var result = PathVisibilityObservation.Read(paths);
            result["Scope"] = "Owner-STA synthetic path attributes/effective-token metadata only; no contents, export, mutation or impersonation.";
            result["AssemblyMvid"] = typeof(VbeSession).Module.ModuleVersionId.ToString("D");
            return result;
        }

        /// <summary>Accepts only version 1 with LocalAppData and Temp as direct GUID-child directories under their expected parents.</summary>
        /// <param name="json">Manifest JSON containing exactly version 1 and LocalAppData/Temp GUID-child directory names.</param>
        /// <param name="localParent">Expected LocalApplicationData parent directory.</param>
        /// <param name="tempParent">Expected system temporary directory.</param>
        /// <returns>Two validated directories and their fixed synthetic-file paths, in observation order.</returns>
        internal static string[] ValidateManifest(string json, string localParent, string tempParent)
        {
            if (json == null || json.Length > 4096) throw new ArgumentException("A bounded synthetic manifest is required.");
            var value = new JavaScriptSerializer().DeserializeObject(json) as IDictionary<string, object>;
            if (value == null || value.Count != 3 || !value.ContainsKey("Version") ||
                !(value["Version"] is int version) || version != 1 || !value.ContainsKey("LocalAppData") || !value.ContainsKey("Temp"))
                throw new ArgumentException("Exactly Version=1, LocalAppData and Temp are required.");
            string local = RequireGuidChild(value["LocalAppData"] as string, localParent);
            string temp = RequireGuidChild(value["Temp"] as string, tempParent);
            if (string.Equals(local, temp, StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException("Two distinct owned GUID directories are required.");
            return new[] { local, Path.Combine(local, SyntheticName), temp, Path.Combine(temp, SyntheticName) };
        }

        /// <summary>Rejects free path fields before any observation is performed.</summary>
        /// <param name="request">Original request JSON, required to contain only the exact diagnostic Command field.</param>
        internal static void RequireParameterFree(string request)
        {
            var value = new JavaScriptSerializer().DeserializeObject(request) as IDictionary<string, object>;
            if (value == null || value.Count != 1 || !value.ContainsKey("Command") ||
                !string.Equals(value["Command"] as string, CommandName, StringComparison.Ordinal))
                throw new ArgumentException("Only the fixed diagnostic Command is permitted.");
        }

        /// <summary>An actual native owner observation must match its captured process and native STA thread.</summary>
        /// <param name="expectedPid">Host process ID captured when this diagnostic was created.</param>
        /// <param name="expectedTid">Native UI-thread ID captured at construction.</param>
        /// <param name="actualPid">Current Win32 process ID.</param>
        /// <param name="actualTid">Current Win32 thread ID.</param>
        /// <param name="apartment">Current managed apartment; only STA is accepted.</param>
        internal static void RequireOwner(int expectedPid, uint expectedTid, uint actualPid, uint actualTid, ApartmentState apartment)
        {
            if (expectedPid <= 0 || expectedTid == 0 || expectedPid != actualPid || expectedTid != actualTid || apartment != ApartmentState.STA)
                throw new InvalidOperationException("The exact owning process/native STA thread is required.");
        }

        /// <summary>Requires a canonical direct child directory whose final component is a GUID and whose ancestors contain no reparse points.</summary>
        /// <param name="directory">Candidate owned synthetic directory.</param>
        /// <param name="parent">Required local parent under which the GUID directory must be placed.</param>
        /// <returns>Canonical full path after all parent, GUID, and reparse checks pass.</returns>
        private static string RequireGuidChild(string directory, string parent)
        {
            ExcelLocalPath(directory); ExcelLocalPath(parent);
            string exact = Path.GetFullPath(directory).TrimEnd(Path.DirectorySeparatorChar);
            Guid guid;
            if (!string.Equals(directory, exact, StringComparison.OrdinalIgnoreCase) ||
                !Guid.TryParseExact(Path.GetFileName(exact), "N", out guid) ||
                !string.Equals(Path.GetDirectoryName(exact), Path.GetFullPath(parent).TrimEnd(Path.DirectorySeparatorChar), StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException("Only a direct GUID child of the specified synthetic parent is permitted.");
            RejectReparseAncestors(exact);
            return exact;
        }

        /// <summary>Rejects non-local, relative, device-prefixed, and alternate-stream paths before filesystem inspection.</summary>
        /// <param name="path">Candidate drive-rooted local filesystem path.</param>
        private static void ExcelLocalPath(string path)
        {
            if (string.IsNullOrWhiteSpace(path) || path.Length < 3 || !char.IsLetter(path[0]) || path[1] != ':' ||
                path[2] != '\\' || path.IndexOf(':', 2) >= 0 || !Path.IsPathRooted(path))
                throw new ArgumentException("An absolute local path without a device prefix or alternate stream is required.");
        }

        /// <summary>Walks the path and its ancestors, refusing existing reparse points that could redirect the diagnostic.</summary>
        /// <param name="path">Validated local path to inspect.</param>
        private static void RejectReparseAncestors(string path)
        {
            for (string current = path; !string.IsNullOrEmpty(current); current = Path.GetDirectoryName(current))
                if ((File.Exists(current) || Directory.Exists(current)) && (File.GetAttributes(current) & FileAttributes.ReparsePoint) != 0)
                    throw new ArgumentException("Diagnostic paths must not traverse filesystem links.");
        }
    }
}
