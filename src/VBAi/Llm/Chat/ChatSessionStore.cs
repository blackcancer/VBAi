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

        /// <summary>Per-process writer identity used to coalesce saves and derive a private recovery filename.</summary>
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
        /// <value>Current read project grants exposed by chat session state.</value>
        public string[] ReadProjectGrants { get; set; } = new string[0];

        /// <summary>Gets or sets the shared context read allowed.</summary>
        /// <value>Current shared context read allowed exposed by chat session state.</value>
        public bool SharedContextReadAllowed { get; set; }

        /// <summary>Gets or sets the read access policy version.</summary>
        /// <value>Current read access policy version exposed by chat session state.</value>
        public int ReadAccessPolicyVersion { get; set; } = 1;

        /// <summary>Gets or sets the provider history start index.</summary>
        /// <value>Current provider history start index exposed by chat session state.</value>
        public int ProviderHistoryStartIndex { get; set; }

        /// <summary>Gets or sets the pending messages.</summary>
        /// <value>Current pending messages exposed by chat session state.</value>
        public List<QueuedChatMessage> PendingMessages { get; set; } = new List<QueuedChatMessage>();

        /// <summary>Gets or sets the budget paused.</summary>
        /// <value>Current budget paused exposed by chat session state.</value>
        public bool BudgetPaused { get; set; }

        /// <summary>Gets or sets the paused turn id.</summary>
        /// <value>Current paused turn id exposed by chat session state.</value>
        public string PausedTurnId { get; set; }

        /// <summary>Gets or sets the paused provider.</summary>
        /// <value>Current paused provider exposed by chat session state.</value>
        public string PausedProvider { get; set; }

        /// <summary>Gets or sets the paused model.</summary>
        /// <value>Current paused model exposed by chat session state.</value>
        public string PausedModel { get; set; }

        /// <summary>Gets or sets the paused effort.</summary>
        /// <value>Current paused effort exposed by chat session state.</value>
        public string PausedEffort { get; set; }

        /// <summary>Gets or sets the paused mode.</summary>
        /// <value>Current paused mode exposed by chat session state.</value>
        public ChatMode PausedMode { get; set; }

        /// <summary>Gets or sets the completed tool actions.</summary>
        /// <value>Current completed tool actions exposed by chat session state.</value>
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

        /// <summary>Empreinte des consignes développeur effectivement appliquées au fil Codex.</summary>
        /// <value>SHA-256 local, ou nul lorsque les consignes du fil ne sont pas connues.</value>
        public string CodexDeveloperInstructionsHash { get; set; }

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
        /// <value>Current draft captured memory exposed by chat session state.</value>
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

    // Uses the SQLite runtime shipped with Windows. SQL values are always bound parameters.
    /// <summary>Persiste les sessions de chat et la mémoire de projet dans la base SQLite Windows.</summary>
    internal sealed partial class ChatSessionStore : IDisposable
    {

        /// <summary>Recognizes temporary conversation scopes that must never be written to SQLite.</summary>
        /// <param name="scope">Conversation scope identifier.</param>
        /// <returns><see langword="true"/> when the scope begins with the reserved <c>temporary:</c> prefix.</returns>
        internal static bool IsTransientScope(string scope) => scope != null && scope.StartsWith("temporary:", StringComparison.Ordinal);

        /// <summary>Handle natif de la base SQLite ouverte.</summary>
        private IntPtr database;

        /// <summary>Gets or sets the database path.</summary>
        /// <value>Current database path exposed by chat session store.</value>
        internal string DatabasePath { get; private set; }

        /// <summary>Sérialiseur JSON configuré pour les charges utiles de session volumineuses.</summary>
        private readonly JavaScriptSerializer json = new JavaScriptSerializer { MaxJsonLength = 32 * 1024 * 1024 };

        /// <summary>Ouvre la base, configure l’attente sur verrou et crée les tables et index nécessaires.</summary>
        /// <param name="path">Chemin du fichier de base SQLite.</param>
        /// <exception cref="IOException">La base ne peut pas être ouverte ou initialisée.</exception>
        public ChatSessionStore(string path) : this(path, false) { }

        /// <summary>Opens the database read-only or read-write/create and initializes schema only for a writable connection.</summary>
        /// <param name="path">SQLite database file path, normalized to a full path.</param>
        /// <param name="readOnly">True to avoid directory creation and schema writes; false to create the database and tables.</param>
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
            if (IsTransientScope(session.Scope)) return;
            string payload = json.Serialize(session);
            session.StorageVersion = SavePayload(session.Id, session.Scope, session.Title, payload, session.StorageVersion);
        }

        /// <summary>Inserts a new conversation or updates its current revision using an optimistic version match.</summary>
        /// <param name="id">Conversation ID used as the SQLite primary key.</param>
        /// <param name="scope">Saved document/project scope; temporary scopes are skipped.</param>
        /// <param name="title">Display title stored with the conversation.</param>
        /// <param name="payload">Serialized chat session snapshot.</param>
        /// <param name="expectedVersion">Previously loaded storage version; null selects insert-only behavior.</param>
        /// <returns>New timestamp-plus-GUID storage version.</returns>
        internal string SavePayload(string id, string scope, string title, string payload, string expectedVersion)
        {
            if (IsTransientScope(scope)) return null;
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

        /// <summary>Deletes only the local conversation revision owned by this writer.</summary>
        /// <param name="id">Conversation ID to delete.</param>
        /// <param name="scope">Saved scope that must match the stored row.</param>
        /// <param name="expectedVersion">Storage revision owned by this writer; deletion is conditional on an exact match.</param>
        internal void Delete(string id, string scope, string expectedVersion)
        {
            if (IsTransientScope(scope)) return;
            if (string.IsNullOrEmpty(id) || string.IsNullOrEmpty(scope)) throw new ArgumentException("A conversation identity and scope are required.");
            Execute("DELETE FROM chat_sessions WHERE id = ?1 AND scope = ?2 AND updated = ?3", id, scope, expectedVersion);
            if (Native.sqlite3_changes(database) == 1) return;
            using (var statement = Prepare("SELECT updated FROM chat_sessions WHERE id = ?1 AND scope = ?2", id, scope))
                if (statement.Step() == 100)
                    throw new IOException(UiText.Get("This conversation changed in another host. Reopen it before deleting."));
        }

        /// <summary>Charge les sessions d’une portée dans l’ordre de mise à jour décroissant.</summary>
        /// <param name="scope">Portée de projet à consulter.</param>
        /// <returns>Sessions désérialisées correspondant à cette portée.</returns>
        public List<ChatSessionState> List(string scope)
        {
            var result = new List<ChatSessionState>();
            if (IsTransientScope(scope)) return result;
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

        /// <summary>Checks whether a saved scope contains at least one chat session.</summary>
        /// <param name="scope">Exact scope key to query.</param>
        /// <returns><see langword="false"/> for transient scopes or when no session row exists.</returns>
        internal bool HasSessions(string scope)
        {
            if (IsTransientScope(scope)) return false;
            using (var statement = Prepare("SELECT 1 FROM chat_sessions WHERE scope = ?1 LIMIT 1", scope)) return statement.Step() == 100;
        }

        /// <summary>Checks for sessions, project memory, or code bookmarks under one saved scope.</summary>
        /// <param name="scope">Exact scope key to query.</param>
        /// <returns><see langword="false"/> for temporary scopes or when all three tables are empty for this scope.</returns>
        internal bool HasScopeData(string scope)
        {
            if (IsTransientScope(scope)) return false;
            using (var statement = Prepare("SELECT 1 FROM chat_sessions WHERE scope = ?1 UNION ALL SELECT 1 FROM project_memory WHERE scope = ?1 UNION ALL SELECT 1 FROM code_bookmarks WHERE scope = ?1 LIMIT 1", scope))
                return statement.Step() == 100;
        }

        /// <summary>Session identity and serialized content staged for atomic promotion into a saved scope.</summary>
        internal sealed class PromotionRow
        {

            /// <summary>Conversation ID, display title, and serialized payload inserted during scope promotion.</summary>
            internal string Id, Title, Payload;
        }

        /// <summary>Internal control-flow signal that promotion found existing data and safely rolled its transaction back.</summary>
        private sealed class ScopeOccupiedException : Exception { }

        /// <summary>Signals that promotion or rollback did not establish a known database outcome; callers must reopen before saving.</summary>
        internal sealed class PromotionOutcomeUnverifiedException : IOException
        {

            /// <summary>Creates an unverified-outcome error while preserving the promotion and optional rollback failures.</summary>
            /// <param name="error">Original promotion failure.</param>
            /// <param name="rollback">Rollback failure, when rollback could not restore a known state.</param>
            internal PromotionOutcomeUnverifiedException(Exception error, Exception rollback) : base(
                "History promotion failed and its database outcome is unverified. Reopen the conversation before saving again.",
                rollback == null ? error : new AggregateException(error, rollback))
            { }
        }

        /// <summary>Claims an empty saved scope and writes its initial conversations and notes atomically.</summary>
        /// <param name="scope">Absolute saved-document scope that must still be empty.</param>
        /// <param name="rows">Conversation snapshots inserted in the promotion transaction.</param>
        /// <param name="memory">Optional project memory written in the same transaction.</param>
        /// <returns>Map from promoted conversation ID to new storage version, or null when the scope is already occupied.</returns>
        internal Dictionary<string, string> PromoteEmptyScope(string scope, IList<PromotionRow> rows, string memory)
        {
            if (IsTransientScope(scope) || !Path.IsPathRooted(scope)) throw new ArgumentException("A saved document scope is required.");
            bool commitAttempted = false;
            try
            {
                Execute("BEGIN IMMEDIATE");
                if (HasScopeData(scope)) throw new ScopeOccupiedException();
                var versions = new Dictionary<string, string>();
                foreach (var row in rows) versions.Add(row.Id, SavePayload(row.Id, scope, row.Title, row.Payload, null));
                if (memory != null) SaveMemory(scope, memory);
                commitAttempted = true; Execute("COMMIT");
                return versions;
            }
            catch (ScopeOccupiedException error)
            {
                if (!RollbackPromotion(out var rollback)) throw new PromotionOutcomeUnverifiedException(error, rollback);
                return null;
            }
            catch (Exception error)
            {
                if (commitAttempted && Native.sqlite3_get_autocommit(database) != 0)
                    throw new PromotionOutcomeUnverifiedException(error, null);
                if (!RollbackPromotion(out var rollback)) throw new PromotionOutcomeUnverifiedException(error, rollback);
                throw;
            }
        }

        /// <summary>Rolls back a promotion transaction and verifies that SQLite returned to autocommit mode.</summary>
        /// <param name="error">Receives the rollback exception, or null when no rollback was needed or it succeeded.</param>
        /// <returns><see langword="true"/> when the transaction is known to be rolled back; otherwise false.</returns>
        private bool RollbackPromotion(out Exception error)
        {
            error = null;
            if (Native.sqlite3_get_autocommit(database) != 0) return true;
            try { Execute("ROLLBACK"); }
            catch (Exception failure) { error = failure; }
            return Native.sqlite3_get_autocommit(database) != 0;
        }

        /// <summary>Worker-produced contents of one scope, published to the UI only after the read connection is disposed.</summary>
        internal sealed class ScopeSnapshot
        {

            /// <summary>Decoded sessions for the requested scope, or null when the caller requested memory only.</summary>
            internal List<ChatSessionState> Sessions;

            /// <summary>Project memory text for the scope; empty when no memory row exists.</summary>
            internal string Memory;
        }

        /// <summary>Reads memory and optionally sessions on a worker-owned read-only SQLite connection.</summary>
        /// <param name="path">Database file to open read-only.</param>
        /// <param name="scope">Exact document/project scope to read.</param>
        /// <param name="includeSessions">True to decode session payloads as well as memory; false returns memory only.</param>
        /// <returns>Task containing detached managed data after the worker connection is disposed.</returns>
        internal static System.Threading.Tasks.Task<ScopeSnapshot> ReadScopeAsync(string path, string scope, bool includeSessions)
        {
            if (IsTransientScope(scope)) return System.Threading.Tasks.Task.FromResult(new ScopeSnapshot { Sessions = new List<ChatSessionState>(), Memory = "" });
            return System.Threading.Tasks.Task.Run(() =>
            {
                using (var store = new ChatSessionStore(path, true))
                    return new ScopeSnapshot
                    {
                        Sessions = includeSessions ? store.List(scope) : null,
                        Memory = store.ReadMemory(scope)
                    };
            });
        }

        /// <summary>Deserializes a stored chat session and restores its persisted queue and state.</summary>
        /// <param name="payload">Serialized session JSON stored in the database.</param>
        /// <returns>Restored session state, with missing policy-version data defaulted to zero, or null for a non-object payload.</returns>
        internal static ChatSessionState DecodeSession(string payload)
        {
            var serializer = new JavaScriptSerializer { MaxJsonLength = 32 * 1024 * 1024 };
            if (!(serializer.DeserializeObject(payload) is Dictionary<string, object> fields)) return null;
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
            if (IsTransientScope(scope)) return "";
            using (var statement = Prepare("SELECT content FROM project_memory WHERE scope = ?1", scope))
                return statement.Step() == 100 ? ReadText(Native.sqlite3_column_text(statement.Handle, 0)) : "";
        }

        /// <summary>Insère ou remplace la mémoire de projet de la portée donnée.</summary>
        /// <param name="scope">Portée de projet à mettre à jour.</param>
        /// <param name="content">Contenu à enregistrer ; une valeur nulle devient une chaîne vide.</param>
        public void SaveMemory(string scope, string content)
        {
            if (IsTransientScope(scope)) return;
            Execute("INSERT OR REPLACE INTO project_memory(scope, content) VALUES (?1, ?2)", scope, content ?? "");
        }

        /// <summary>Prépare une instruction SQLite et lie ses valeurs comme paramètres texte UTF-8.</summary>
        /// <param name="sql">Instruction SQL à préparer.</param>
        /// <param name="values">Valeurs des paramètres numérotés de l’instruction.</param>
        /// <returns>État natif prêt à être exécuté et libéré par <see cref="Statement.Dispose()"/>.</returns>
        private Statement Prepare(string sql, params string[] values)
        {
            Check(Native.sqlite3_prepare_v2(database, Utf8(sql), -1, out IntPtr handle, IntPtr.Zero));
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

            /// <summary>Checks whether the connection is outside any active transaction.</summary>
            /// <param name="db">SQLite connection handle.</param>
            /// <returns>Nonzero when SQLite is in autocommit mode; zero while a transaction remains active.</returns>
            [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl)] internal static extern int sqlite3_get_autocommit(IntPtr db);

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
            /// <param name="db">SQLite connection whose most recent statement is inspected.</param>
            /// <returns>Number of rows inserted, updated, or deleted by the most recent write.</returns>
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
