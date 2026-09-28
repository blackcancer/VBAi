using System;
using System.Diagnostics;
using System.Reflection;
using System.Text;

namespace CodexVBE
{
    /// <summary>Démarre un processus avec une entrée redirigée UTF-8 sans préambule BOM.</summary>
    internal static class ProcessInput
    {
        /// <summary>Sérialise temporairement la modification du réglage global d’encodage d’entrée.</summary>
        private static readonly object startLock = new object();
        /// <summary>Résout le cache d’encodage du runtime .NET Framework installé.</summary>
        internal static Func<FieldInfo> InputEncodingField = () => typeof(Console).GetField("_inputEncoding", BindingFlags.NonPublic | BindingFlags.Static);

        /// <summary>Remplace temporairement l’encodage d’entrée de Console afin que .NET Framework ne préfixe pas le flux binaire d’un BOM.</summary>
        /// <param name="process">Processus configuré avec une entrée standard redirigée.</param>
        /// <returns>Valeur renvoyée par Process.Start.</returns>
        /// <exception cref="InvalidOperationException">Le runtime ne permet pas de retrouver son encodage d’entrée interne.</exception>
        internal static bool StartWithoutPreamble(Process process)
        {
            // .NET Framework has no ProcessStartInfo.StandardInputEncoding. Its redirected
            // StreamWriter takes Console.InputEncoding and can emit a BOM before binary data.
            // The public Console setter can fail in GUI hosts. Scope and serialize the cached
            // value override across all our protocol clients; restore it even when Start fails.
            lock (startLock)
            {
                var field = InputEncodingField();
                if (field == null) throw new InvalidOperationException("Impossible de configurer une entrée binaire sans préambule sur ce runtime .NET Framework.");
                object previous = field.GetValue(null);
                try { field.SetValue(null, new UTF8Encoding(false)); return process.Start(); }
                finally { field.SetValue(null, previous); }
            }
        }
    }
}
