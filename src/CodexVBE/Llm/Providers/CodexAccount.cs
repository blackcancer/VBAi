using System;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;

namespace CodexVBE
{
    /// <summary>Résume l’état de connexion au compte ChatGPT détecté par le client Codex.</summary>
    internal sealed class CodexAccountStatus
    {
        /// <summary>Indique si la session CLI est connectée via ChatGPT.</summary>
        /// <value>Vrai lorsque la sortie de statut confirme une connexion ChatGPT.</value>
        public bool ChatGptConnected { get; private set; }
        /// <summary>Texte localisé à présenter pour cet état de compte.</summary>
        /// <value>Statut de connexion ou détail d’échec.</value>
        public string Text { get; private set; }
        /// <summary>Crée le statut retourné par la vérification du compte.</summary>
        /// <param name="connected">Indique si le compte ChatGPT est connecté.</param>
        /// <param name="text">Texte de statut à afficher.</param>
        public CodexAccountStatus(bool connected, string text) { ChatGptConnected = connected; Text = text; }
    }

    /// <summary>Interroge le client Codex installé et ouvre sa procédure de connexion.</summary>
    internal static class CodexAccount
    {
        /// <summary>Démarre le client natif configuré, y compris son écran de connexion.</summary>
        internal static Func<ProcessStartInfo, Process> StartProcess = Process.Start;
        /// <summary>Attend la vérification native avec son délai de dix secondes.</summary>
        internal static Func<Process, int, bool> WaitForExit = (process, milliseconds) => process.WaitForExit(milliseconds);
        /// <summary>Vérifie la présence du client à son emplacement installé.</summary>
        internal static Func<string, bool> FileExists = File.Exists;
        /// <summary>Résout le chemin du client Codex depuis sa configuration ou son emplacement usuel.</summary>
        /// <value>Chemin configuré, chemin installé, ou « codex.exe » si aucun fichier connu ne le confirme.</value>
        public static string Executable
        {
            get
            {
                string configured = Environment.GetEnvironmentVariable("CODEXVBE_CODEX_CLI");
                if (!string.IsNullOrWhiteSpace(configured)) return configured;
                string installed = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "Programs", "OpenAI", "Codex", "bin", "codex.exe");
                return FileExists(installed) ? installed : "codex.exe";
            }
        }

        /// <summary>Exécute « codex login status » et interprète le résultat pour déterminer la connexion ChatGPT.</summary>
        /// <returns>Tâche qui fournit le statut de connexion ; elle échoue si la commande dépasse dix secondes.</returns>
        public static Task<CodexAccountStatus> ReadStatusAsync()
        {
            return Task.Run(() => {
                var info = new ProcessStartInfo(Executable, "login status") {
                    UseShellExecute = false, CreateNoWindow = true, RedirectStandardOutput = true,
                    RedirectStandardError = true
                };
                using (var process = StartProcess(info))
                {
                    var output = process.StandardOutput.ReadToEndAsync();
                    var error = process.StandardError.ReadToEndAsync();
                    if (!WaitForExit(process, 10000)) { process.Kill(); throw new TimeoutException(UiText.Get("Codex verification timed out.")); }
                    // Drain both pipes concurrently so neither can block the deadline check.
                    Task.WaitAll(output, error);
                    string result = (output.Result + " " + error.Result).Trim();
                    if (process.ExitCode != 0)
                        return new CodexAccountStatus(false, UiText.Get("Not signed in to ChatGPT") +
                            (result.Length == 0 ? "." : " : " + result));
                    bool chatGpt = result.IndexOf("ChatGPT", StringComparison.OrdinalIgnoreCase) >= 0;
                    return new CodexAccountStatus(chatGpt,
                        chatGpt ? UiText.Get("Connected to ChatGPT") : UiText.Get("Codex is connected, but not through ChatGPT: ") + result);
                }
            });
        }

        /// <summary>Ouvre la commande interactive de connexion du client Codex.</summary>
        public static void StartLogin()
        {
            StartProcess(new ProcessStartInfo(Executable, "login") { UseShellExecute = true });
        }
    }
}
