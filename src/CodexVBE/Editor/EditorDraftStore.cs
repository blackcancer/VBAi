using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Web.Script.Serialization;
using System.Linq;

namespace CodexVBE
{
    internal sealed class EditorDraft
    {
        public string Baseline { get; set; }
        public string Text { get; set; }
        public string Key { get; set; }
    }
    /// <summary>Private recovery drafts, encrypted for the current Windows account; never overwrite another process's draft.</summary>
    internal sealed class EditorDraftStore
    {
        internal readonly string Root;
        private readonly object gate = new object();
        private readonly string owner = System.Diagnostics.Process.GetCurrentProcess().Id + "-" + Guid.NewGuid().ToString("N");
        private DateTime lastCleanup;
        private readonly JavaScriptSerializer json = new JavaScriptSerializer { MaxJsonLength = 16 * 1024 * 1024 };
        internal EditorDraftStore(string root = null) { Root = root ?? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "CodexVBE", "EditorDrafts"); }
        private string DirectoryFor(string key) => Path.Combine(Root, EditorDocument.Hash(key));
        internal void Save(EditorDocument document)
        {
            SaveSnapshot(document.Id, new EditorDraft { Key = document.RecoveryKey, Baseline = document.Baseline, Text = document.Text });
        }
        internal void SaveSnapshot(string id, EditorDraft data)
        { lock (gate) SaveSnapshotCore(id, data); }
        private void SaveSnapshotCore(string id, EditorDraft data)
        {
            if (DateTime.UtcNow - lastCleanup > TimeSpan.FromDays(1)) { Cleanup(DateTime.UtcNow); lastCleanup = DateTime.UtcNow; }
            if (data.Text == data.Baseline) return;
            string directory = DirectoryFor(data.Key); Directory.CreateDirectory(directory);
            string destination = Path.Combine(directory, owner + "-" + id + ".draft");
            string temporary = destination + ".tmp";
            byte[] bytes = ProtectedData.Protect(Encoding.UTF8.GetBytes(json.Serialize(data)), null, DataProtectionScope.CurrentUser);
            try { File.WriteAllBytes(temporary, bytes); if (File.Exists(destination)) File.Replace(temporary, destination, null); else File.Move(temporary, destination); }
            finally { if (File.Exists(temporary)) File.Delete(temporary); }
        }
        internal EditorDraft Recover(string key)
        { lock (gate) return RecoverCore(key); }
        private EditorDraft RecoverCore(string key)
        {
            string directory = DirectoryFor(key);
            if (!Directory.Exists(directory)) return null;
            var files = new DirectoryInfo(directory).GetFiles("*.draft"); Array.Sort(files, (a, b) => b.LastWriteTimeUtc.CompareTo(a.LastWriteTimeUtc));
            foreach (var file in files)
                try
                {
                    if (file.Length > 16 * 1024 * 1024) continue;
                    var draft = json.Deserialize<EditorDraft>(Encoding.UTF8.GetString(ProtectedData.Unprotect(File.ReadAllBytes(file.FullName), null, DataProtectionScope.CurrentUser)));
                    if (draft.Key != key) continue;
                    EditorDocument.Validate(draft.Baseline); EditorDocument.Validate(draft.Text); return draft;
                }
                catch (Exception) { } // A corrupt draft cannot prevent opening the current source.
            return null;
        }
        internal void ClearOwn(EditorDocument document)
        { lock (gate) { string path = Path.Combine(DirectoryFor(document.RecoveryKey), owner + "-" + document.Id + ".draft"); if (File.Exists(path)) File.Delete(path); } }
        internal void Cleanup(DateTime now)
        {
            lock (gate)
            {
                if (!Directory.Exists(Root) || (File.GetAttributes(Root) & FileAttributes.ReparsePoint) != 0) return;
                foreach (var directory in new DirectoryInfo(Root).GetDirectories())
                {
                    if ((directory.Attributes & FileAttributes.ReparsePoint) != 0) continue;
                    // Keep the newest recovery file for each module, even beyond retention.
                    var files = directory.GetFiles("*.draft").OrderByDescending(f => f.LastWriteTimeUtc).ToArray();
                    foreach (var file in files.Skip(1))
                    {
                        if (file.LastWriteTimeUtc >= now.AddDays(-30) || (file.Attributes & FileAttributes.ReparsePoint) != 0) continue;
                        if (!int.TryParse(file.Name.Split('-')[0], out int pid)) continue;
                        try { using (var process = System.Diagnostics.Process.GetProcessById(pid)) { if (!process.HasExited) continue; } }
                        catch (ArgumentException) { }
                        try { file.Delete(); } catch (IOException) { } catch (UnauthorizedAccessException) { }
                    }
                }
            }
        }
    }
}
