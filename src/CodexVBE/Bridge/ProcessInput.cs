using System;
using System.Diagnostics;
using System.Reflection;
using System.Text;

namespace CodexVBE
{
    internal static class ProcessInput
    {
        private static readonly object startLock = new object();

        internal static bool StartWithoutPreamble(Process process)
        {
            // .NET Framework has no ProcessStartInfo.StandardInputEncoding. Its redirected
            // StreamWriter takes Console.InputEncoding and can emit a BOM before binary data.
            // The public Console setter can fail in GUI hosts. Scope and serialize the cached
            // value override across all our protocol clients; restore it even when Start fails.
            lock (startLock)
            {
                var field = typeof(Console).GetField("_inputEncoding", BindingFlags.NonPublic | BindingFlags.Static);
                if (field == null) throw new InvalidOperationException("Impossible de configurer une entrée binaire sans préambule sur ce runtime .NET Framework.");
                object previous = field.GetValue(null);
                try { field.SetValue(null, new UTF8Encoding(false)); return process.Start(); }
                finally { field.SetValue(null, previous); }
            }
        }
    }
}
