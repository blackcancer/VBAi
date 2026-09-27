using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Web.Script.Serialization;

namespace CodexVBE
{
    internal sealed class ChatEntry
    {
        public string Speaker { get; set; }
        public string Text { get; set; }
        public string StreamId { get; set; }
        public CodeChange Change { get; set; }
        public VbeChatReference[] References { get; set; }
        public string AttachedMemory { get; set; }
    }

    internal sealed class ChatSessionState
    {
        public string Id { get; set; } = Guid.NewGuid().ToString("N");
        public string Scope { get; set; }
        public string Title { get; set; } = "Nouvelle conversation";
        public bool Archived { get; set; }
        public string Provider { get; set; } = "Codex";
        public string Model { get; set; }
        public string Effort { get; set; }
        public string CodexThreadId { get; set; }
        public string MessagesJson { get; set; }
        public string Draft { get; set; }
        public VbeChatReference[] DraftReferences { get; set; }
        public List<ChatEntry> Entries { get; set; } = new List<ChatEntry>();
        public override string ToString() { return Title; }
    }

    // Uses the SQLite runtime shipped with Windows. SQL values are always bound parameters.
    internal sealed class ChatSessionStore : IDisposable
    {
        private IntPtr database;
        private readonly JavaScriptSerializer json = new JavaScriptSerializer { MaxJsonLength = 32 * 1024 * 1024 };
        public ChatSessionStore(string path)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(path)));
            int result = Native.sqlite3_open_v2(Utf8(path), out database, 6, IntPtr.Zero);
            if (result != 0) { Dispose(); throw new IOException("Impossible d'ouvrir l'historique SQLite."); }
            try
            {
                Native.sqlite3_busy_timeout(database, 1500);
                Execute("CREATE TABLE IF NOT EXISTS chat_sessions (id TEXT PRIMARY KEY, scope TEXT NOT NULL, title TEXT NOT NULL, updated TEXT NOT NULL, payload TEXT NOT NULL)");
                Execute("CREATE INDEX IF NOT EXISTS chat_sessions_scope ON chat_sessions(scope, updated)");
                Execute("CREATE TABLE IF NOT EXISTS project_memory (scope TEXT PRIMARY KEY, content TEXT NOT NULL)");
            }
            catch { Dispose(); throw; }
        }

        public void Save(ChatSessionState session)
        {
            Execute("INSERT OR REPLACE INTO chat_sessions(id, scope, title, updated, payload) VALUES (?1, ?2, ?3, ?4, ?5)",
                session.Id, session.Scope, session.Title, DateTime.UtcNow.ToString("O"), json.Serialize(session));
        }

        public List<ChatSessionState> List(string scope)
        {
            var result = new List<ChatSessionState>();
            using (var statement = Prepare("SELECT payload FROM chat_sessions WHERE scope = ?1 ORDER BY updated DESC", scope))
            {
                while (statement.Step() == 100)
                {
                    var session = json.Deserialize<ChatSessionState>(ReadText(Native.sqlite3_column_text(statement.Handle, 0)));
                    if (session != null && session.Scope == scope) result.Add(session);
                }
            }
            return result;
        }

        private void Execute(string sql, params string[] values)
        {
            using (var statement = Prepare(sql, values)) statement.Step();
        }

        public string ReadMemory(string scope)
        {
            using (var statement = Prepare("SELECT content FROM project_memory WHERE scope = ?1", scope))
                return statement.Step() == 100 ? ReadText(Native.sqlite3_column_text(statement.Handle, 0)) : "";
        }

        public void SaveMemory(string scope, string content)
        {
            Execute("INSERT OR REPLACE INTO project_memory(scope, content) VALUES (?1, ?2)", scope, content ?? "");
        }

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

        private void Check(int result)
        {
            if (result != 0 && result != 100 && result != 101)
                throw new IOException("Historique SQLite : " + ReadText(Native.sqlite3_errmsg(database)));
        }

        private static byte[] Utf8(string text) { return Encoding.UTF8.GetBytes(text + "\0"); }
        private static string ReadText(IntPtr ptr)
        {
            if (ptr == IntPtr.Zero) return "";
            int count = 0;
            while (Marshal.ReadByte(ptr, count) != 0) count++;
            byte[] bytes = new byte[count];
            Marshal.Copy(ptr, bytes, 0, count);
            return Encoding.UTF8.GetString(bytes);
        }

        public void Dispose()
        {
            if (database != IntPtr.Zero) { Native.sqlite3_close_v2(database); database = IntPtr.Zero; }
        }

        private sealed class Statement : IDisposable
        {
            private readonly ChatSessionStore owner;
            public IntPtr Handle { get; private set; }
            public Statement(ChatSessionStore owner, IntPtr handle) { this.owner = owner; Handle = handle; }
            public int Step() { int result = Native.sqlite3_step(Handle); owner.Check(result); return result; }
            public void Dispose() { if (Handle != IntPtr.Zero) { Native.sqlite3_finalize(Handle); Handle = IntPtr.Zero; } }
        }

        private static class Native
        {
            [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl)] internal static extern int sqlite3_open_v2(byte[] path, out IntPtr db, int flags, IntPtr vfs);
            [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl)] internal static extern int sqlite3_close_v2(IntPtr db);
            [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl)] internal static extern int sqlite3_busy_timeout(IntPtr db, int ms);
            [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl)] internal static extern int sqlite3_prepare_v2(IntPtr db, byte[] sql, int count, out IntPtr statement, IntPtr tail);
            [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl)] internal static extern int sqlite3_bind_text(IntPtr statement, int index, byte[] text, int count, IntPtr destructor);
            [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl)] internal static extern int sqlite3_step(IntPtr statement);
            [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl)] internal static extern int sqlite3_finalize(IntPtr statement);
            [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl)] internal static extern IntPtr sqlite3_column_text(IntPtr statement, int column);
            [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl)] internal static extern IntPtr sqlite3_errmsg(IntPtr db);
        }
    }
}
