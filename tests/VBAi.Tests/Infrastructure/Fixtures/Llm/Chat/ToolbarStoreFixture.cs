namespace VBAi.Tests.Unit
{
    using System;
    using System.Runtime.InteropServices;
    using VBAi;

    public sealed partial class ChatSessionStoreTests
    {
        private static VbeToolbarProfiles.Bar StoredBar(string name = "VBAi - One") => new VbeToolbarProfiles.Bar { Name = name, Position = 1, Commands = new VbeToolbarProfiles.Command[0] };
        // Observe le SQL réellement préparé, pour échouer une seule étape du moteur natif.
        [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl)]
        private static extern IntPtr sqlite3_sql(IntPtr statement);
        private static string ToolbarStatementSql(IntPtr statement) => Marshal.PtrToStringAnsi(sqlite3_sql(statement));
    }
}
