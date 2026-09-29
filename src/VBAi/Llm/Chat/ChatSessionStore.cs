using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Web.Script.Serialization;

namespace VBAi
{
    /// <summary>Enregistre un message affiché ou généré dans une session de conversation.</summary>
    internal sealed class ChatEntry
    {
        /// <summary>Obtient ou définit l’identifiant du locuteur, par exemple utilisateur ou assistant.</summary>
        /// <value>Nom du locuteur.</value>
        public string Speaker { get; set; }
        /// <summary>Étape native détaillée lorsqu’il s’agit d’une activité de l’agent.</summary>
        /// <value>Métadonnées de l’action, ou null pour un ancien message.</value>
        public CodexAgentActivity Activity { get; set; }
        /// <summary>Obtient ou définit le texte du message.</summary>
        /// <value>Texte brut du message.</value>
        public string Text { get; set; }
        /// <summary>Obtient ou définit l’identifiant du flux auquel le message appartient, s’il est diffusé.</summary>
        /// <value>Identifiant du flux ou nul.</value>
        public string StreamId { get; set; }
        /// <summary>Obtient ou définit le changement de code associé au message, le cas échéant.</summary>
        /// <value>Changement de code ou nul.</value>
        public CodeChange Change { get; set; }
        /// <summary>Coupe de concepteur proposant une récupération du formulaire.</summary>
        /// <value>Changement récupérable ou null.</value>
        public FormCutChange FormCut { get; set; }
        /// <summary>Obtient ou définit les références VBE citées par le message.</summary>
        /// <value>Références du message.</value>
        public VbeChatReference[] References { get; set; }
        /// <summary>Obtient ou définit la mémoire de projet jointe au message.</summary>
        /// <value>Texte de mémoire ou nul.</value>
        public string AttachedMemory { get; set; }
        /// <summary>Obtient ou définit les pièces jointes du message.</summary>
        /// <value>Pièces jointes de l’entrée.</value>
        public ChatAttachment[] Attachments { get; set; }
        /// <summary>Obtient ou définit l’identifiant du tour de conversation associé.</summary>
        /// <value>Identifiant du tour courant.</value>
        public string TurnId { get; set; }
    }

    /// <summary>Contient l’état persistant d’une conversation, son brouillon et ses entrées.</summary>
    internal sealed class ChatSessionState
    {
        /// <summary>Opaque database revision loaded with this snapshot; never sent to a provider.</summary>
        internal string StorageVersion;
        internal readonly string WriterId = Guid.NewGuid().ToString("N");
        /// <summary>Obtient ou définit l’identifiant stable de la session.</summary>
        /// <value>Identifiant texte au format GUID compact.</value>
        public string Id { get; set; } = Guid.NewGuid().ToString("N");
        /// <summary>Obtient ou définit la portée de projet à laquelle appartient la session.</summary>
        /// <value>Clé de portée du projet.</value>
        public string Scope { get; set; }
        /// <summary>Obtient ou définit le titre persistant de la session.</summary>
        /// <value>Titre stocké.</value>
        public string Title { get; set; } = "Nouvelle conversation";
        /// <summary>Obtient ou définit l’état archivé de la session.</summary>
        /// <value><see langword="true"/> si la session est archivée.</value>
        public bool Archived { get; set; }
        /// <summary>Obtient ou définit l’état épinglé de la session.</summary>
        /// <value><see langword="true"/> si la session est épinglée.</value>
        public bool Pinned { get; set; }
        /// <summary>Obtient ou définit le mode de conversation.</summary>
        /// <value>Mode choisi.</value>
        public ChatMode Mode { get; set; } = ChatMode.Agent;
        /// <summary>Gets or sets the read project grants.</summary>
        /// <value>The current value represented by this member.</value>
        public string[] ReadProjectGrants { get; set; } = new string[0];
        /// <summary>Gets or sets the shared context read allowed.</summary>
        /// <value>The current value represented by this member.</value>
        public bool SharedContextReadAllowed { get; set; }
        /// <summary>Gets or sets the read access policy version.</summary>
        /// <value>The current value represented by this member.</value>
        public int ReadAccessPolicyVersion { get; set; } = 1;
        /// <summary>Gets or sets the provider history start index.</summary>
        /// <value>The current value represented by this member.</value>
        public int ProviderHistoryStartIndex { get; set; }
        /// <summary>Gets or sets the pending messages.</summary>
        /// <value>The current value represented by this member.</value>
        public List<QueuedChatMessage> PendingMessages { get; set; } = new List<QueuedChatMessage>();
        /// <summary>Gets or sets the budget paused.</summary>
        /// <value>The current value represented by this member.</value>
        public bool BudgetPaused { get; set; }
        /// <summary>Gets or sets the paused turn id.</summary>
        /// <value>The current value represented by this member.</value>
        public string PausedTurnId { get; set; }
        /// <summary>Gets or sets the paused provider.</summary>
        /// <value>The current value represented by this member.</value>
        public string PausedProvider { get; set; }
        /// <summary>Gets or sets the paused model.</summary>
        /// <value>The current value represented by this member.</value>
        public string PausedModel { get; set; }
        /// <summary>Gets or sets the paused effort.</summary>
        /// <value>The current value represented by this member.</value>
        public string PausedEffort { get; set; }
        /// <summary>Gets or sets the paused mode.</summary>
        /// <value>The current value represented by this member.</value>
        public ChatMode PausedMode { get; set; }
        /// <summary>Gets or sets the completed tool actions.</summary>
        /// <value>The current value represented by this member.</value>
        public List<string> CompletedToolActions { get; set; } = new List<string>();
        /// <summary>Obtient ou définit les pièces jointes du brouillon.</summary>
        /// <value>Pièces jointes en attente du prochain message.</value>
        public ChatAttachment[] DraftAttachments { get; set; }
        /// <summary>Obtient ou définit le fournisseur utilisé pour la session.</summary>
        /// <value>Nom du fournisseur.</value>
        public string Provider { get; set; } = "Codex";
        /// <summary>Obtient ou définit l’identifiant du modèle choisi.</summary>
        /// <value>Identifiant de modèle ou nul.</value>
        public string Model { get; set; }
        /// <summary>Obtient ou définit le niveau d’effort choisi pour le modèle.</summary>
        /// <value>Identifiant d’effort ou nul.</value>
        public string Effort { get; set; }
        /// <summary>Obtient ou définit l’identifiant de fil de conversation Codex associé.</summary>
        /// <value>Identifiant de fil, ou nul si non applicable.</value>
        public string CodexThreadId { get; set; }
        /// <summary>Dossier privé auquel appartient le fil ; nul pour un ancien fil extérieur.</summary>
        /// <value>Chemin du stockage Codex ayant créé ce fil.</value>
        public string CodexThreadHome { get; set; }
        /// <summary>Obtient ou définit le contexte nécessaire pour reprendre la conversation.</summary>
        /// <value>Contexte de reprise ou nul.</value>
        public string ResumeContext { get; set; }
        /// <summary>Obtient ou définit la représentation JSON des messages conservée pour compatibilité.</summary>
        /// <value>Messages sérialisés, ou nul.</value>
        public string MessagesJson { get; set; }
        /// <summary>Obtient ou définit le texte du brouillon courant.</summary>
        /// <value>Texte du compositeur.</value>
        public string Draft { get; set; }
        /// <summary>Gets or sets the draft captured memory.</summary>
        /// <value>The current value represented by this member.</value>
        public string DraftCapturedMemory { get; set; }
        /// <summary>Obtient ou définit les références VBE du brouillon.</summary>
        /// <value>Références sélectionnées pour le prochain message.</value>
        public VbeChatReference[] DraftReferences { get; set; }
        /// <summary>Obtient ou définit les entrées de la session.</summary>
        /// <value>Messages et changements ordonnés de la conversation.</value>
        public List<ChatEntry> Entries { get; set; } = new List<ChatEntry>();
        /// <summary>Obtient le titre destiné à l’affichage, traduit lorsque le titre est celui par défaut.</summary>
        /// <value>Titre localisé ou titre personnalisé.</value>
        [ScriptIgnore]
        public string DisplayTitle { get { return Title == "Nouvelle conversation" ? UiText.Get("New conversation") : Title; } }
        /// <summary>Retourne le titre d’affichage précédé d’une étoile lorsque la session est épinglée.</summary>
        /// <returns>Le libellé présenté dans la liste des sessions.</returns>
        public override string ToString() { return (Pinned ? "★ " : "") + DisplayTitle; }
    }

    /// <summary>Persiste les sessions de chat et la mémoire de projet dans la base SQLite Windows.</summary>
    // Uses the SQLite runtime shipped with Windows. SQL values are always bound parameters.
    internal sealed partial class ChatSessionStore : IDisposable
    {
        /// <summary>Handle natif de la base SQLite ouverte.</summary>
        private IntPtr database;
        internal string DatabasePath { get; private set; }
        /// <summary>Sérialiseur JSON configuré pour les charges utiles de session volumineuses.</summary>
        private readonly JavaScriptSerializer json = new JavaScriptSerializer { MaxJsonLength = 32 * 1024 * 1024 };
        /// <summary>Ouvre la base, configure l’attente sur verrou et crée les tables et index nécessaires.</summary>
        /// <param name="path">Chemin du fichier de base SQLite.</param>
        /// <exception cref="IOException">La base ne peut pas être ouverte ou initialisée.</exception>
        public ChatSessionStore(string path) : this(path, false) { }

        private ChatSessionStore(string path, bool readOnly)
        {
            DatabasePath = Path.GetFullPath(path);
            if (!readOnly) Directory.CreateDirectory(Path.GetDirectoryName(DatabasePath));
            // SQLITE_OPEN_READONLY versus SQLITE_OPEN_READWRITE | SQLITE_OPEN_CREATE.
            int result = Native.sqlite3_open_v2(Utf8(DatabasePath), out database, readOnly ? 1 : 6, IntPtr.Zero);
            if (result != 0) { Dispose(); throw new IOException(UiText.Get("Unable to open SQLite history.")); }
            try
            {
                Native.sqlite3_busy_timeout(database, 1500);
                if (readOnly) return;
                Execute("CREATE TABLE IF NOT EXISTS chat_sessions (id TEXT PRIMARY KEY, scope TEXT NOT NULL, title TEXT NOT NULL, updated TEXT NOT NULL, payload TEXT NOT NULL)");
                Execute("CREATE INDEX IF NOT EXISTS chat_sessions_scope ON chat_sessions(scope, updated)");
                Execute("CREATE TABLE IF NOT EXISTS code_bookmarks (scope TEXT NOT NULL, name_key TEXT NOT NULL, payload TEXT NOT NULL, PRIMARY KEY(scope, name_key))");
                Execute("CREATE TABLE IF NOT EXISTS project_memory (scope TEXT PRIMARY KEY, content TEXT NOT NULL)");
            }
            catch { Dispose(); throw; }
        }

        /// <summary>Insère ou remplace la session sérialisée avec son horodatage UTC.</summary>
        /// <param name="session">État de session à persister.</param>
        public void Save(ChatSessionState session)
        {
            string payload = json.Serialize(session);
            session.StorageVersion = SavePayload(session.Id, session.Scope, session.Title, payload, session.StorageVersion);
        }

        /// <summary>Writes an immutable UI snapshot on the connection's owning worker.</summary>
        internal string SavePayload(string id, string scope, string title, string payload, string expectedVersion)
        {
            string version = DateTime.UtcNow.ToString("O") + "-" + Guid.NewGuid().ToString("N");
            if (expectedVersion == null)
                Execute("INSERT OR IGNORE INTO chat_sessions(id, scope, title, updated, payload) VALUES (?1, ?2, ?3, ?4, ?5)",
                    id, scope, title, version, payload);
            else
                Execute("UPDATE chat_sessions SET title = ?3, updated = ?4, payload = ?5 WHERE id = ?1 AND scope = ?2 AND updated = ?6",
                    id, scope, title, version, payload, expectedVersion);
            if (Native.sqlite3_changes(database) != 1)
                throw new IOException("This conversation was changed in another host. Your local draft is retained; reopen the conversation before saving again.");
            return version;
        }

        /// <summary>Charge les sessions d’une portée dans l’ordre de mise à jour décroissant.</summary>
        /// <param name="scope">Portée de projet à consulter.</param>
        /// <returns>Sessions désérialisées correspondant à cette portée.</returns>
        public List<ChatSessionState> List(string scope)
        {
            var result = new List<ChatSessionState>();
            using (var statement = Prepare("SELECT payload, updated FROM chat_sessions WHERE scope = ?1 ORDER BY updated DESC", scope))
            {
                while (statement.Step() == 100)
                {
                    var session = DecodeSession(ReadText(Native.sqlite3_column_text(statement.Handle, 0)));
                    if (session != null && session.Scope == scope)
                    {
                        session.StorageVersion = ReadText(Native.sqlite3_column_text(statement.Handle, 1));
                        result.Add(session);
                    }
                }
            }
            return result;
        }

        internal sealed class ScopeSnapshot
        {
            internal List<ChatSessionState> Sessions;
            internal string Memory;
        }

        /// <summary>Owns the read connection and decoded objects entirely on a worker until publication.</summary>
        internal static System.Threading.Tasks.Task<ScopeSnapshot> ReadScopeAsync(string path, string scope, bool includeSessions)
        {
            return System.Threading.Tasks.Task.Run(() => {
                using (var store = new ChatSessionStore(path, true))
                    return new ScopeSnapshot {
                        Sessions = includeSessions ? store.List(scope) : null,
                        Memory = store.ReadMemory(scope)
                    };
            });
        }

        /// <summary>Deserializes a stored chat session and restores its persisted queue and state.</summary>
        /// <param name="payload">Text containing the payload.</param>
        /// <returns>The result produced by this operation.</returns>
        internal static ChatSessionState DecodeSession(string payload)
        {
            var serializer = new JavaScriptSerializer { MaxJsonLength = 32 * 1024 * 1024 };
            var fields = serializer.DeserializeObject(payload) as Dictionary<string, object>;
            if (fields == null) return null;
            var session = serializer.ConvertToType<ChatSessionState>(fields);
            if (!fields.ContainsKey(nameof(ChatSessionState.ReadAccessPolicyVersion))) session.ReadAccessPolicyVersion = 0;
            return session;
        }

        /// <summary>Prépare, exécute une instruction paramétrée puis libère son état natif.</summary>
        /// <param name="sql">Instruction SQLite à exécuter.</param>
        /// <param name="values">Valeurs liées dans l’ordre des paramètres SQL numérotés.</param>
        private void Execute(string sql, params string[] values)
        {
            using (var statement = Prepare(sql, values)) statement.Step();
        }

        /// <summary>Lit la mémoire de projet associée à une portée.</summary>
        /// <param name="scope">Portée de projet recherchée.</param>
        /// <returns>Contenu enregistré, ou une chaîne vide si aucune mémoire n’existe.</returns>
        public string ReadMemory(string scope)
        {
            using (var statement = Prepare("SELECT content FROM project_memory WHERE scope = ?1", scope))
                return statement.Step() == 100 ? ReadText(Native.sqlite3_column_text(statement.Handle, 0)) : "";
        }

        /// <summary>Insère ou remplace la mémoire de projet de la portée donnée.</summary>
        /// <param name="scope">Portée de projet à mettre à jour.</param>
        /// <param name="content">Contenu à enregistrer ; une valeur nulle devient une chaîne vide.</param>
        public void SaveMemory(string scope, string content)
        {
            Execute("INSERT OR REPLACE INTO project_memory(scope, content) VALUES (?1, ?2)", scope, content ?? "");
        }

        /// <summary>Prépare une instruction SQLite et lie ses valeurs comme paramètres texte UTF-8.</summary>
        /// <param name="sql">Instruction SQL à préparer.</param>
        /// <param name="values">Valeurs des paramètres numérotés de l’instruction.</param>
        /// <returns>État natif prêt à être exécuté et libéré par <see cref="Statement.Dispose()"/>.</returns>
        private Statement Prepare(string sql, params string[] values)
        {
            IntPtr handle;
            Check(Native.sqlite3_prepare_v2(database, Utf8(sql), -1, out handle, IntPtr.Zero));
            var statement = new Statement(this, handle);
            try
            {
                for (int i = 0; i < values.Length; i++)
                {
                    byte[] bytes = Utf8(values[i] ?? "");
                    Check(Native.sqlite3_bind_text(handle, i + 1, bytes, bytes.Length - 1, new IntPtr(-1)));
                }
                return statement;
            }
            catch { statement.Dispose(); throw; }
        }

        /// <summary>Vérifie un code de retour SQLite et convertit les erreurs en exception d’E/S.</summary>
        /// <param name="result">Code renvoyé par SQLite.</param>
        /// <exception cref="IOException">SQLite a renvoyé un code autre que succès, ligne ou terminé.</exception>
        private void Check(int result)
        {
            if (result != 0 && result != 100 && result != 101)
                throw new IOException("Historique SQLite : " + ReadText(Native.sqlite3_errmsg(database)));
        }

        /// <summary>Encode un texte en UTF-8 terminé par un octet nul pour l’API native.</summary>
        /// <param name="text">Texte à encoder.</param>
        /// <returns>Octets UTF-8 terminés par zéro.</returns>
        private static byte[] Utf8(string text) { return Encoding.UTF8.GetBytes(text + "\0"); }
        /// <summary>Lit une chaîne UTF-8 terminée par zéro depuis un pointeur natif.</summary>
        /// <param name="ptr">Pointeur vers les octets natifs, éventuellement nul.</param>
        /// <returns>Texte décodé, ou une chaîne vide si le pointeur est nul.</returns>
        private static string ReadText(IntPtr ptr)
        {
            if (ptr == IntPtr.Zero) return "";
            int count = 0;
            while (Marshal.ReadByte(ptr, count) != 0) count++;
            byte[] bytes = new byte[count];
            Marshal.Copy(ptr, bytes, 0, count);
            return Encoding.UTF8.GetString(bytes);
        }

        /// <summary>Ferme la base SQLite et remet son handle à zéro.</summary>
        public void Dispose()
        {
            if (database != IntPtr.Zero) { Native.sqlite3_close_v2(database); database = IntPtr.Zero; }
        }

        /// <summary>Possède un état d’instruction SQLite préparée et le finalise à sa libération.</summary>
        private sealed class Statement : IDisposable
        {
            /// <summary>Base propriétaire utilisée pour vérifier les erreurs d’exécution.</summary>
            private readonly ChatSessionStore owner;
            /// <summary>Obtient le handle de l’instruction native préparée.</summary>
            /// <value>Handle SQLite de l’instruction, remis à zéro après finalisation.</value>
            public IntPtr Handle { get; private set; }
            /// <summary>Associe le handle d’instruction à son magasin propriétaire.</summary>
            /// <param name="owner">Magasin qui fournit la vérification des erreurs SQLite.</param>
            /// <param name="handle">Handle de l’instruction préparée.</param>
            public Statement(ChatSessionStore owner, IntPtr handle) { this.owner = owner; Handle = handle; }
            /// <summary>Exécute une étape de l’instruction et vérifie son code de retour.</summary>
            /// <returns>Code SQLite de l’étape, notamment ligne disponible ou fin des résultats.</returns>
            public int Step() { int result = owner.StepNative(Handle); owner.Check(result); return result; }
            /// <summary>Finalise l’instruction native une seule fois.</summary>
            public void Dispose() { if (Handle != IntPtr.Zero) { Native.sqlite3_finalize(Handle); Handle = IntPtr.Zero; } }
        }

        /// <summary>Déclarations P/Invoke des fonctions utilisées dans winsqlite3.dll.</summary>
        private static class Native
        {
            /// <summary>Ouvre une base SQLite à partir d’un chemin encodé en UTF-8.</summary>
            /// <param name="path">Chemin terminé par zéro du fichier de base.</param>
            /// <param name="db">Reçoit le handle de base, y compris en cas d’échec partiel.</param>
            /// <param name="flags">Options d’ouverture SQLite.</param>
            /// <param name="vfs">VFS SQLite à utiliser, ou nul pour le VFS par défaut.</param>
            /// <returns>Code de résultat SQLite.</returns>
            [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl)] internal static extern int sqlite3_open_v2(byte[] path, out IntPtr db, int flags, IntPtr vfs);
            /// <summary>Ferme une base SQLite et libère ses ressources.</summary>
            /// <param name="db">Handle de la base à fermer.</param>
            /// <returns>Code de résultat SQLite.</returns>
            [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl)] internal static extern int sqlite3_close_v2(IntPtr db);
            /// <summary>Définit le délai maximal d’attente d’un verrou SQLite.</summary>
            /// <param name="db">Handle de la base.</param>
            /// <param name="ms">Durée d’attente en millisecondes.</param>
            /// <returns>Code de résultat SQLite.</returns>
            [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl)] internal static extern int sqlite3_busy_timeout(IntPtr db, int ms);
            /// <summary>Compile une instruction SQL en instruction préparée.</summary>
            /// <param name="db">Handle de la base.</param>
            /// <param name="sql">Instruction encodée en UTF-8.</param>
            /// <param name="count">Nombre maximal d’octets SQL à analyser.</param>
            /// <param name="statement">Reçoit le handle de l’instruction préparée.</param>
            /// <param name="tail">Pointeur facultatif vers le reste du texte SQL.</param>
            /// <returns>Code de résultat SQLite.</returns>
            [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl)] internal static extern int sqlite3_prepare_v2(IntPtr db, byte[] sql, int count, out IntPtr statement, IntPtr tail);
            /// <summary>Liaison d’une valeur texte à un paramètre d’instruction.</summary>
            /// <param name="statement">Instruction préparée.</param>
            /// <param name="index">Index du paramètre, à partir de un.</param>
            /// <param name="text">Octets UTF-8 de la valeur.</param>
            /// <param name="count">Nombre d’octets à lire.</param>
            /// <param name="destructor">Destructeur SQLite de la mémoire ; -1 indique que SQLite peut la copier.</param>
            /// <returns>Code de résultat SQLite.</returns>
            [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl)] internal static extern int sqlite3_bind_text(IntPtr statement, int index, byte[] text, int count, IntPtr destructor);
            /// <summary>Exécute l’étape suivante d’une instruction préparée.</summary>
            /// <param name="statement">Instruction préparée.</param>
            /// <returns>Code indiquant une ligne, la fin des résultats ou une erreur.</returns>
            [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl)] internal static extern int sqlite3_step(IntPtr statement);
            /// <summary>Returns the number of rows changed by the most recent write on this connection.</summary>
            [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl)] internal static extern int sqlite3_changes(IntPtr db);
            /// <summary>Finalise une instruction préparée et libère ses ressources.</summary>
            /// <param name="statement">Instruction à finaliser.</param>
            /// <returns>Code de résultat SQLite.</returns>
            [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl)] internal static extern int sqlite3_finalize(IntPtr statement);
            /// <summary>Obtient le pointeur vers la valeur texte d’une colonne de la ligne courante.</summary>
            /// <param name="statement">Instruction ayant produit la ligne courante.</param>
            /// <param name="column">Index de la colonne, à partir de zéro.</param>
            /// <returns>Pointeur natif vers la valeur UTF-8 ou nul.</returns>
            [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl)] internal static extern IntPtr sqlite3_column_text(IntPtr statement, int column);
            /// <summary>Obtient le message d’erreur associé à une base SQLite.</summary>
            /// <param name="db">Handle de la base concernée.</param>
            /// <returns>Pointeur vers le message UTF-8 géré par SQLite.</returns>
            [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl)] internal static extern IntPtr sqlite3_errmsg(IntPtr db);
        }
    }
}
