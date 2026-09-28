using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Web.Script.Serialization;
using System.Linq;

namespace CodexVBE
{
    /// <summary>Charge utile persistée pour restaurer un brouillon VBA après fermeture ou incident.</summary>
internal sealed class EditorDraft
    {
        /// <summary>Source native servant de base au brouillon.</summary>
        /// <value>Version de départ utilisée pour détecter les conflits au rétablissement.</value>
public string Baseline { get; set; }
        /// <summary>Source modifiée à restaurer dans l’éditeur.</summary>
        /// <value>Texte courant du brouillon.</value>
public string Text { get; set; }
        /// <summary>Clé du projet et du composant auquel le brouillon appartient.</summary>
        /// <value>Clé de récupération de l’adaptateur VBE.</value>
public string Key { get; set; }
    }
    /// <summary>Private recovery drafts, encrypted for the current Windows account; never overwrite another process's draft.</summary>
    internal sealed class EditorDraftStore
    {
        /// <summary>Répertoire racine des brouillons chiffrés.</summary>
internal readonly string Root;
        /// <summary>Protège les opérations de lecture, écriture et nettoyage de cette instance.</summary>
private readonly object gate = new object();
        /// <summary>Préfixe distinguant les fichiers créés par cette instance du magasin.</summary>
private readonly string owner = System.Diagnostics.Process.GetCurrentProcess().Id + "-" + Guid.NewGuid().ToString("N");
        /// <summary>Heure du dernier nettoyage périodique effectué par cette instance.</summary>
private DateTime lastCleanup;
        /// <summary>Stores the read process used by EditorDraftStore.</summary>
internal Func<int, System.Diagnostics.Process> ReadProcess = System.Diagnostics.Process.GetProcessById;
        /// <summary>Stores the read attributes used by EditorDraftStore.</summary>
internal Func<FileSystemInfo, FileAttributes> ReadAttributes = NativeAttributes;
        /// <summary>Performs the native attributes operation for EditorDraftStore.</summary>
/// <param name="item">The item used by this operation.</param>
/// <returns>The result produced by this operation.</returns>
private static FileAttributes NativeAttributes(FileSystemInfo item) => item.Attributes;
        /// <summary>Sérialiseur JSON des instantanés de brouillon.</summary>
private readonly JavaScriptSerializer json = new JavaScriptSerializer { MaxJsonLength = 16 * 1024 * 1024 };
        /// <summary>Crée le magasin dans le chemin fourni ou dans le dossier local de données de l’utilisateur.</summary>
        /// <param name="root">Répertoire facultatif de stockage, principalement utile pour isoler le magasin.</param>
internal EditorDraftStore(string root = null) { Root = root ?? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "CodexVBE", "EditorDrafts"); }
        /// <summary>Calcule le sous-répertoire associé à la clé d’un module.</summary>
        /// <param name="key">Clé stable de récupération du module.</param>
        /// <returns>Chemin combinant la racine et l’empreinte de la clé.</returns>
private string DirectoryFor(string key) => Path.Combine(Root, EditorDocument.Hash(key));
        /// <summary>Enregistre la base et le texte courant du document.</summary>
        /// <param name="document">Document dont le brouillon doit être protégé.</param>
internal void Save(EditorDocument document)
        {
            SaveSnapshot(document.Id, new EditorDraft { Key = document.RecoveryKey, Baseline = document.Baseline, Text = document.Text });
        }
        /// <summary>Enregistre un instantané sous un identifiant de document, de manière synchronisée.</summary>
        /// <param name="id">Identifiant propre au document dans la session courante.</param>
        /// <param name="data">Données de récupération à chiffrer.</param>
internal void SaveSnapshot(string id, EditorDraft data)
        { lock (gate) SaveSnapshotCore(id, data); }
        /// <summary>Écrit atomiquement l’instantané chiffré et déclenche au besoin le nettoyage périodique.</summary>
        /// <param name="id">Identifiant du document.</param>
        /// <param name="data">Données de récupération.</param>
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
        /// <summary>Recherche et déchiffre le brouillon valide le plus récent pour une clé de module.</summary>
        /// <param name="key">Clé stable du projet et du composant.</param>
        /// <returns>Brouillon récupérable, ou <see langword="null"/> si aucun fichier valide n’existe.</returns>
internal EditorDraft Recover(string key)
        { lock (gate) return RecoverCore(key); }
        /// <summary>Parcourt les instantanés de la clé et ignore les fichiers corrompus ou trop volumineux.</summary>
        /// <param name="key">Clé dont le brouillon doit être restauré.</param>
        /// <returns>Premier instantané valide, en commençant par le plus récent, ou <see langword="null"/>.</returns>
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
        /// <summary>Supprime le fichier de brouillon appartenant à cette instance pour le document indiqué.</summary>
        /// <param name="document">Document dont l’instantané local doit être effacé.</param>
internal void ClearOwn(EditorDocument document)
        { lock (gate) { string path = Path.Combine(DirectoryFor(document.RecoveryKey), owner + "-" + document.Id + ".draft"); if (File.Exists(path)) File.Delete(path); } }
        /// <summary>Supprime les brouillons anciens de processus arrêtés en conservant le plus récent de chaque module.</summary>
        /// <param name="now">Heure UTC utilisée pour appliquer la durée de conservation.</param>
internal void Cleanup(DateTime now)
        {
            lock (gate)
            {
                if (!Directory.Exists(Root) || (File.GetAttributes(Root) & FileAttributes.ReparsePoint) != 0) return;
                foreach (var directory in new DirectoryInfo(Root).GetDirectories())
                {
                    if ((ReadAttributes(directory) & FileAttributes.ReparsePoint) != 0) continue;
                    // Keep the newest recovery file for each module, even beyond retention.
                    var files = directory.GetFiles("*.draft").OrderByDescending(f => f.LastWriteTimeUtc).ToArray();
                    foreach (var file in files.Skip(1))
                    {
                        if (file.LastWriteTimeUtc >= now.AddDays(-30) || (ReadAttributes(file) & FileAttributes.ReparsePoint) != 0) continue;
                        if (!int.TryParse(file.Name.Split('-')[0], out int pid)) continue;
                        try { using (var process = ReadProcess(pid)) { if (!process.HasExited) continue; } }
                        catch (ArgumentException) { }
                        try { file.Delete(); } catch (IOException) { } catch (UnauthorizedAccessException) { }
                    }
                }
            }
        }
    }
}
