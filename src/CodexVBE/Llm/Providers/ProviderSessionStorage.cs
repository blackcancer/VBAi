using System;
using System.Diagnostics;
using System.IO;

namespace CodexVBE
{
    /// <summary>Isole l’état des clients CLI sous les données locales du complément.</summary>
    internal static class ProviderSessionStorage
    {
        /// <summary>Dossier Codex privé, indépendant de CODEX_HOME hérité et de Codex Desktop.</summary>
        internal static string CodexHome => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "CodexVBE", "Providers", "Codex");
        /// <summary>Dossier Copilot privé, indépendant de COPILOT_HOME hérité.</summary>
        internal static string CopilotHome => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "CodexVBE", "Providers", "Copilot");
        /// <summary>Configure exclusivement l’environnement du processus enfant Codex.</summary>
        /// <param name="info">Processus à lancer sans interpréteur shell.</param>
        internal static void ConfigureCodex(ProcessStartInfo info) { Directory.CreateDirectory(CodexHome); info.EnvironmentVariables["CODEX_HOME"] = CodexHome; }
        /// <summary>Configure exclusivement l’environnement du processus enfant Copilot.</summary>
        /// <param name="info">Processus à lancer sans interpréteur shell.</param>
        internal static void ConfigureCopilot(ProcessStartInfo info) { Directory.CreateDirectory(CopilotHome); info.EnvironmentVariables["COPILOT_HOME"] = CopilotHome; }
        /// <summary>Convertit une reprise CLI extérieure en contexte local, sans lire ni modifier son stockage.</summary>
        /// <param name="session">Conversation locale à migrer.</param>
        internal static void PrepareCodexSession(ChatSessionState session)
        {
            if (session == null || string.IsNullOrEmpty(session.CodexThreadId) ||
                string.Equals(session.CodexThreadHome, CodexHome, StringComparison.OrdinalIgnoreCase)) return;
            session.ResumeContext = session.ResumeContext + "\n\n" + ChatHistory.Export(session);
            session.CodexThreadId = null;
            session.CodexThreadHome = null;
        }
    }
}
