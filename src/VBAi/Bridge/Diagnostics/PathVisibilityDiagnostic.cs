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
        internal const string EnvironmentName = "VBAi_TEST_PATH_VISIBILITY_MANIFEST";
        internal const string CommandName = "diagnostic_path_visibility";
        internal const string SyntheticName = "VBAi.PathVisibility.synthetic";
        private readonly int ownerPid;
        private readonly uint ownerTid;
        private readonly string[] paths;
        private readonly string preparationError;

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

        internal IDictionary<string, object> Read()
        {
            RequireOwner(ownerPid, ownerTid, PathVisibilityObservation.ProcessId, PathVisibilityObservation.ThreadId, Thread.CurrentThread.GetApartmentState());
            if (paths == null) throw new InvalidOperationException(preparationError == null ? "Path visibility diagnostic is disabled." : "Invalid diagnostic manifest: " + preparationError);
            var result = PathVisibilityObservation.Read(paths);
            result["Scope"] = "Owner-STA synthetic path attributes/effective-token metadata only; no contents, export, mutation or impersonation.";
            result["AssemblyMvid"] = typeof(VbeSession).Module.ModuleVersionId.ToString("D");
            return result;
        }

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
        internal static void RequireParameterFree(string request)
        {
            var value = new JavaScriptSerializer().DeserializeObject(request) as IDictionary<string, object>;
            if (value == null || value.Count != 1 || !value.ContainsKey("Command") ||
                !string.Equals(value["Command"] as string, CommandName, StringComparison.Ordinal))
                throw new ArgumentException("Only the fixed diagnostic Command is permitted.");
        }

        /// <summary>An actual native owner observation must match its captured process and native STA thread.</summary>
        internal static void RequireOwner(int expectedPid, uint expectedTid, uint actualPid, uint actualTid, ApartmentState apartment)
        {
            if (expectedPid <= 0 || expectedTid == 0 || expectedPid != actualPid || expectedTid != actualTid || apartment != ApartmentState.STA)
                throw new InvalidOperationException("The exact owning process/native STA thread is required.");
        }

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
        private static void ExcelLocalPath(string path)
        {
            if (string.IsNullOrWhiteSpace(path) || path.Length < 3 || !char.IsLetter(path[0]) || path[1] != ':' ||
                path[2] != '\\' || path.IndexOf(':', 2) >= 0 || !Path.IsPathRooted(path))
                throw new ArgumentException("An absolute local path without a device prefix or alternate stream is required.");
        }
        private static void RejectReparseAncestors(string path)
        {
            for (string current = path; !string.IsNullOrEmpty(current); current = Path.GetDirectoryName(current))
                if ((File.Exists(current) || Directory.Exists(current)) && (File.GetAttributes(current) & FileAttributes.ReparsePoint) != 0)
                    throw new ArgumentException("Diagnostic paths must not traverse filesystem links.");
        }
    }
}
