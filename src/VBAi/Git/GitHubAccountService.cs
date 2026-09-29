using System;
using System.Diagnostics;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace VBAi
{
    /// <summary>Liste et authentifie les comptes GitHub gérés par Git Credential Manager.</summary>
    internal sealed class GitHubAccountService
    {
        /// <summary>Exécuteur des commandes GCM, injectable pour fournir un transport alternatif.</summary>
        private readonly Func<string, CancellationToken, Task<string>> execute;
        /// <summary>Démarre le processus Git natif sans modifier le protocole GCM.</summary>
        internal Action<Process> StartProcess = process => process.Start();
        /// <summary>Attend la fin du processus avec le délai de production.</summary>
        internal Func<Process, int, bool> WaitForExit = (process, milliseconds) => process.WaitForExit(milliseconds);
        /// <summary>Vérifie si le processus est déjà terminé avant l’annulation.</summary>
        internal Func<Process, bool> HasExited = process => process.HasExited;
        /// <summary>Interrompt uniquement le processus démarré par cette opération.</summary>
        internal Action<Process> KillProcess = process => process.Kill();
        /// <summary>Crée le service avec l’exécuteur fourni, ou l’exécution système par défaut.</summary>
        /// <param name="execute">Fonction d’exécution facultative recevant les arguments et le jeton d’annulation.</param>
        internal GitHubAccountService(Func<string, CancellationToken, Task<string>> execute = null)
        { this.execute = execute ?? Execute; }

        /// <summary>Récupère et valide les noms des comptes GitHub présents dans GCM.</summary>
        /// <param name="cancellation">Jeton d’annulation transmis à la commande.</param>
        /// <returns>Comptes distincts, triés sans tenir compte de la casse.</returns>
        /// <exception cref="InvalidOperationException">La sortie GCM contient une entrée qui n’est pas un nom de compte valide.</exception>
        internal async Task<string[]> ListAsync(CancellationToken cancellation)
        {
            return ParseAccounts(await execute("credential-manager github list --url https://github.com --no-ui", cancellation));
        }
        /// <summary>Lance la connexion GitHub dans le navigateur au moyen de GCM.</summary>
        /// <param name="cancellation">Jeton d’annulation transmis à la commande.</param>
        /// <returns>Tâche terminée lorsque GCM a fini le processus de connexion.</returns>
        internal async Task LoginAsync(CancellationToken cancellation)
        {
            await execute("credential-manager github login --url https://github.com --browser", cancellation);
        }
        /// <summary>Vérifie qu’une chaîne correspond au format d’un nom de compte GitHub.</summary>
        /// <param name="value">Nom à contrôler.</param>
        /// <returns><see langword="true"/> si le nom satisfait le motif de compte, sinon <see langword="false"/>.</returns>
        internal static bool ValidAccount(string value)
        {
            return Regex.IsMatch(value ?? "", @"^[A-Za-z0-9](?:[A-Za-z0-9-]{0,37}[A-Za-z0-9])?$");
        }
        /// <summary>Analyse la liste de comptes GCM, valide chaque ligne et retourne des valeurs distinctes triées.</summary>
        /// <param name="text">Sortie texte de la commande de liste.</param>
        /// <returns>Comptes nettoyés, distincts sans tenir compte de la casse et triés.</returns>
        /// <exception cref="InvalidOperationException">Une ligne ne contient pas un nom de compte reconnu.</exception>
        internal static string[] ParseAccounts(string text)
        {
            var lines = (text ?? "").Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries).Select(x => x.Trim()).ToArray();
            if (lines.Any(x => !ValidAccount(x)))
                throw new InvalidOperationException(UiText.Get("Unexpected Git Credential Manager response. Check its version."));
            return lines.Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(x => x, StringComparer.OrdinalIgnoreCase).ToArray();
        }
        /// <summary>Exécute git.exe sans fenêtre, capture ses flux et annule ou interrompt une commande trop longue.</summary>
        /// <param name="arguments">Arguments transmis à git.exe.</param>
        /// <param name="cancellation">Jeton d’annulation de l’opération.</param>
        /// <returns>Sortie standard de la commande si elle se termine avec le code zéro.</returns>
        /// <exception cref="InvalidOperationException">Git est absent ou GCM termine avec une erreur.</exception>
        /// <exception cref="TimeoutException">La commande dépasse cinq minutes.</exception>
        /// <exception cref="OperationCanceledException">L’annulation est demandée.</exception>
        private Task<string> Execute(string arguments, CancellationToken cancellation)
        {
            return Task.Run(async () => {
                cancellation.ThrowIfCancellationRequested();
                var start = new ProcessStartInfo("git.exe", arguments) {
                    UseShellExecute = false, CreateNoWindow = true, RedirectStandardOutput = true, RedirectStandardError = true
                };
                using (var process = new Process { StartInfo = start })
                {
                    try { StartProcess(process); }
                    catch (System.ComponentModel.Win32Exception) { throw new InvalidOperationException(UiText.Get("Git for Windows is required. Install it with Git Credential Manager, then reopen settings.")); }
                    using (cancellation.Register(() => { try { if (!HasExited(process)) KillProcess(process); } catch (InvalidOperationException) { } catch (System.ComponentModel.Win32Exception) { } }))
                    {
                        var output = process.StandardOutput.ReadToEndAsync();
                        var error = process.StandardError.ReadToEndAsync();
                        if (!WaitForExit(process, 300000)) { try { KillProcess(process); } catch { } throw new TimeoutException(UiText.Get("GitHub sign-in timed out. Start sign-in again from settings.")); }
                        await Task.WhenAll(output, error);
                        cancellation.ThrowIfCancellationRequested();
                        // Do not expose raw authentication output or diagnostic bodies in the UI.
                        if (process.ExitCode != 0) throw new InvalidOperationException(UiText.Get("Git Credential Manager did not finish the operation. Check its installation or sign in to GitHub again."));
                        return output.Result;
                    }
                }
            }, cancellation);
        }
    }
}
