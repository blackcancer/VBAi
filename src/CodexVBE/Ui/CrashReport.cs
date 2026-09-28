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
        internal const string Repository = "https://github.com/blackcancer/CodexVBE";
        internal const string Recipient = "init-sys-rev@hotmail.com";
        internal string Id { get; } = Guid.NewGuid().ToString("N");
        internal static string DirectoryPath => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "CodexVBE", "CrashReports");
        internal string TechnicalDetails { get; }
        internal string DefaultTitle { get; }

        internal CrashReport(Exception error = null)
        {
            var assembly = typeof(CrashReport).Assembly;
            var text = new StringBuilder();
            text.AppendLine("Report: " + Id);
            text.AppendLine("UTC: " + DateTime.UtcNow.ToString("O"));
            text.AppendLine("VBAi: " + (assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion ?? assembly.GetName().Version.ToString()));
            using (var process = Process.GetCurrentProcess()) text.AppendLine("Host: " + process.ProcessName);
            text.AppendLine("Platform: Windows " + (Environment.Is64BitProcess ? "x64" : "x86"));
            text.AppendLine("CLR: " + Environment.Version);
            text.AppendLine("Interface: " + UiText.Culture.Name);
            text.AppendLine("Theme: " + UiTheme.Choice);
            int count = 0;
            for (var current = error; current != null && count++ < 5; current = current.InnerException)
            {
                text.AppendLine("Exception: " + current.GetType().FullName);
                var frames = new StackTrace(current, false).GetFrames();
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

        internal string Body(string title, string description)
        {
            if (string.IsNullOrWhiteSpace(title) || title.Length > 180) throw new ArgumentException("Invalid report title.");
            if ((description?.Length ?? 0) > 8000) throw new ArgumentException("Report description is too long.");
            string body = "# " + title.Replace("\r", " ").Replace("\n", " ") + "\n\n" + (description ?? "") +
                "\n\n## Technical details\n\n```text\n" + TechnicalDetails + "```\n";
            return body.Replace("\r\n", "\n").Replace("\n", "\r\n");
        }

        internal string Save(string body, string directory = null)
        {
            directory = directory ?? DirectoryPath;
            Directory.CreateDirectory(directory);
            string path = Path.Combine(directory, Id + ".md");
            File.WriteAllText(path, body, new UTF8Encoding(false));
            return path;
        }

        private CrashReport(Snapshot snapshot)
        {
            Id = snapshot.Id;
            TechnicalDetails = snapshot.Details;
            DefaultTitle = snapshot.Title;
        }
        internal sealed class Snapshot
        {
            public string Id { get; set; }
            public string Details { get; set; }
            public string Title { get; set; }
        }
        internal void SavePending(string directory = null)
        {
            directory = directory ?? DirectoryPath;
            Save(Body(DefaultTitle, ""), directory);
            File.WriteAllText(Path.Combine(directory, Id + ".pending"),
                new JavaScriptSerializer().Serialize(new Snapshot { Id = Id, Details = TechnicalDetails, Title = DefaultTitle }), new UTF8Encoding(false));
        }
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

        internal static bool IsOwned(Exception error)
        {
            if (error is AggregateException aggregate)
                foreach (var inner in aggregate.InnerExceptions) if (IsOwned(inner)) return true;
            for (var current = error; current != null; current = current.InnerException)
            {
                var frames = new StackTrace(current, false).GetFrames();
                if (frames != null) foreach (var frame in frames)
                    if (frame.GetMethod()?.DeclaringType?.Assembly == typeof(CrashReport).Assembly) return true;
            }
            return false;
        }
    }
}
