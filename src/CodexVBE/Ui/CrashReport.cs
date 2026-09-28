using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Text;
using System.Web.Script.Serialization;

namespace CodexVBE
{
    /// <summary>Rapport technique volontairement limité aux métadonnées et aux méthodes, sans message brut ni chemin.</summary>
    internal sealed class CrashReport
    {
        /// <summary>Stores the metadata assembly used by CrashReport.</summary>
internal static Func<Assembly> MetadataAssembly = ReadMetadataAssembly;
        /// <summary>Stores the process is64 bit used by CrashReport.</summary>
internal static Func<bool> ProcessIs64Bit = ReadProcessIs64Bit;
        /// <summary>Stores the runtime version used by CrashReport.</summary>
internal static Func<Version> RuntimeVersion = ReadRuntimeVersion;
        /// <summary>Performs the read runtime version operation for CrashReport.</summary>
/// <returns>The result produced by this operation.</returns>
private static Version ReadRuntimeVersion() => Environment.Version;
        /// <summary>Stores the frame snapshot used by CrashReport.</summary>
internal static Func<Exception, StackFrame[]> FrameSnapshot = ReadFrames;
        /// <summary>Performs the read metadata assembly operation for CrashReport.</summary>
/// <returns>The result produced by this operation.</returns>
private static Assembly ReadMetadataAssembly() => typeof(CrashReport).Assembly;
        /// <summary>Performs the read process is64 bit operation for CrashReport.</summary>
/// <returns>The result produced by this operation.</returns>
private static bool ReadProcessIs64Bit() => Environment.Is64BitProcess;
        /// <summary>Performs the read frames operation for CrashReport.</summary>
/// <param name="error">The error used by this operation.</param>
/// <returns>The result produced by this operation.</returns>
private static StackFrame[] ReadFrames(Exception error) => new StackTrace(error, false).GetFrames();
        /// <summary>GitHub repository used as the destination for product issue reports.</summary>
internal const string Repository = "https://github.com/blackcancer/CodexVBE";
        /// <summary>Support mailbox used by Outlook and local mail drafts.</summary>
internal const string Recipient = "init-sys-rev@hotmail.com";
        /// <summary>Gets the unique identifier assigned to this report.</summary><value>32-character lowercase GUID without separators.</value>
internal string Id { get; } = Guid.NewGuid().ToString("N");
        /// <summary>Stores the report directory used by CrashReport.</summary>
internal static Func<string> ReportDirectory = NativeReportDirectory;
        /// <summary>Performs the native report directory operation for CrashReport.</summary>
/// <returns>The result produced by this operation.</returns>
private static string NativeReportDirectory() => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "CodexVBE", "CrashReports");
        /// <summary>Gets the per-user directory used to save crash reports.</summary><value>Local application-data CrashReports directory.</value>
internal static string DirectoryPath => ReportDirectory();
        /// <summary>Gets the technical metadata and stack method names captured for the report.</summary><value>Sanitized technical report text.</value>
internal string TechnicalDetails { get; }
        /// <summary>Gets the initial issue title based on the supplied exception or generic report action.</summary><value>Default editable issue title.</value>
internal string DefaultTitle { get; }

        /// <summary>Captures product, host, platform, interface, theme, and a bounded exception method trace.</summary>
        /// <param name="error">Optional exception whose types and method names are included.</param>
internal CrashReport(Exception error = null)
        {
            var assembly = MetadataAssembly();
            var text = new StringBuilder();
            text.AppendLine("Report: " + Id);
            text.AppendLine("UTC: " + DateTime.UtcNow.ToString("O"));
            text.AppendLine("VBAi: " + (assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion ?? assembly.GetName().Version.ToString()));
            using (var process = Process.GetCurrentProcess()) text.AppendLine("Host: " + process.ProcessName);
            text.AppendLine("Platform: Windows " + (ProcessIs64Bit() ? "x64" : "x86"));
            text.AppendLine("CLR: " + RuntimeVersion());
            text.AppendLine("Interface: " + UiText.Culture.Name);
            text.AppendLine("Theme: " + UiTheme.Choice);
            int count = 0;
            for (var current = error; current != null && count++ < 5; current = current.InnerException)
            {
                text.AppendLine("Exception: " + current.GetType().FullName);
                var frames = FrameSnapshot(current);
                if (frames == null) continue;
                for (int i = 0; i < Math.Min(frames.Length, 40); i++)
                {
                    var method = frames[i].GetMethod();
                    // No Exception.Message, ToString(), source file, arguments, VBA, log or settings.
                    if (method != null) text.AppendLine("  at " + method.DeclaringType?.FullName + "." + method.Name);
                }
            }
            TechnicalDetails = text.ToString();
            DefaultTitle = error == null ? UiText.Get("VBAi · Report an issue") : "VBAi · " + error.GetType().Name;
        }

        /// <summary>Formats a Markdown issue body with the title, user description, and captured technical details.</summary>
        /// <param name="title">Issue title, limited to 180 characters.</param><param name="description">User description, limited to 8,000 characters.</param>
        /// <returns>Markdown with Windows line endings.</returns>
        /// <exception cref="ArgumentException">The title is empty or too long, or the description exceeds its limit.</exception>
internal string Body(string title, string description)
        {
            if (string.IsNullOrWhiteSpace(title) || title.Length > 180) throw new ArgumentException("Invalid report title.");
            if ((description?.Length ?? 0) > 8000) throw new ArgumentException("Report description is too long.");
            string body = "# " + title.Replace("\r", " ").Replace("\n", " ") + "\n\n" + (description ?? "") +
                "\n\n## Technical details\n\n```text\n" + TechnicalDetails + "```\n";
            return body.Replace("\r\n", "\n").Replace("\n", "\r\n");
        }

        /// <summary>Saves the report body to a new Markdown file named with this report's identifier.</summary>
        /// <param name="body">Markdown body to write.</param><param name="directory">Destination directory, or null to use <see cref="DirectoryPath"/>.</param>
        /// <returns>Full path of the saved report.</returns>
internal string Save(string body, string directory = null)
        {
            directory = directory ?? DirectoryPath;
            Directory.CreateDirectory(directory);
            string path = Path.Combine(directory, Id + ".md");
            File.WriteAllText(path, body, new UTF8Encoding(false));
            return path;
        }

        /// <summary>Reconstructs a report from a validated pending-report snapshot.</summary>
        /// <param name="snapshot">Snapshot containing its identifier, details, and title.</param>
private CrashReport(Snapshot snapshot)
        {
            Id = snapshot.Id;
            TechnicalDetails = snapshot.Details;
            DefaultTitle = snapshot.Title;
        }
        /// <summary>Serialized data retained so a fatal report can be offered after the next start.</summary>
internal sealed class Snapshot
        {
            /// <summary>Report identifier used to locate its Markdown backup.</summary>
            /// <value>Unique report ID.</value>
public string Id { get; set; }
            /// <summary>Sanitized technical detail text.</summary>
            /// <value>Technical report body.</value>
public string Details { get; set; }
            /// <summary>Title displayed when the pending report is reopened.</summary>
            /// <value>Initial report title.</value>
public string Title { get; set; }
        }
        /// <summary>Saves the report and writes a marker for presentation after a later startup.</summary>
        /// <param name="directory">Destination directory, or null to use the local report directory.</param>
internal void SavePending(string directory = null)
        {
            directory = directory ?? DirectoryPath;
            Save(Body(DefaultTitle, ""), directory);
            File.WriteAllText(Path.Combine(directory, Id + ".pending"),
                new JavaScriptSerializer().Serialize(new Snapshot { Id = Id, Details = TechnicalDetails, Title = DefaultTitle }), new UTF8Encoding(false));
        }
        /// <summary>Loads a pending report marker and reconstructs its saved report metadata.</summary>
        /// <param name="path">Path to the pending JSON marker.</param><returns>Reconstructed report.</returns>
        /// <exception cref="InvalidDataException">The marker is oversized, malformed, or inconsistent with its file name.</exception>
internal static CrashReport ReadPending(string path)
        {
            if (new FileInfo(path).Length > 50000) throw new InvalidDataException("Invalid report size.");
            var snapshot = new JavaScriptSerializer().Deserialize<Snapshot>(File.ReadAllText(path));
            if (snapshot == null || !Guid.TryParseExact(snapshot.Id, "N", out var id) ||
                snapshot.Id != Path.GetFileNameWithoutExtension(path) || string.IsNullOrEmpty(snapshot.Details) ||
                string.IsNullOrWhiteSpace(snapshot.Title) || snapshot.Title.Length > 180)
                throw new InvalidDataException("Invalid report snapshot.");
            return new CrashReport(snapshot);
        }

        /// <summary>Checks whether an exception stack contains a method from this add-in assembly.</summary>
        /// <param name="error">Exception and inner-exception chain to inspect.</param><returns>Whether an add-in method appears in the stack.</returns>
internal static bool IsOwned(Exception error)
        {
            if (error is AggregateException aggregate)
                foreach (var inner in aggregate.InnerExceptions) if (IsOwned(inner)) return true;
            for (var current = error; current != null; current = current.InnerException)
            {
                var frames = FrameSnapshot(current);
                if (frames != null) foreach (var frame in frames)
                    if (frame.GetMethod()?.DeclaringType?.Assembly == typeof(CrashReport).Assembly) return true;
            }
            return false;
        }
    }
}
