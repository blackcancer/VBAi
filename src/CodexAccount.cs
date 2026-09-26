using System;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;

namespace CodexVBE
{
    internal static class CodexAccount
    {
        public static string Executable
        {
            get
            {
                string configured = Environment.GetEnvironmentVariable("CODEXVBE_CODEX_CLI");
                if (!string.IsNullOrWhiteSpace(configured)) return configured;
                string installed = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "Programs", "OpenAI", "Codex", "bin", "codex.exe");
                return File.Exists(installed) ? installed : "codex.exe";
            }
        }

        public static Task<string> ReadStatusAsync()
        {
            return Task.Run(() => {
                var info = new ProcessStartInfo(Executable, "login status") {
                    UseShellExecute = false, CreateNoWindow = true, RedirectStandardOutput = true,
                    RedirectStandardError = true
                };
                using (var process = Process.Start(info))
                {
                    string output = process.StandardOutput.ReadToEnd();
                    string error = process.StandardError.ReadToEnd();
                    if (!process.WaitForExit(10000)) { process.Kill(); throw new TimeoutException("Vérification Codex trop longue."); }
                    if (process.ExitCode != 0) return "ChatGPT non connecté : " + error.Trim();
                    return output.Trim();
                }
            });
        }

        public static void StartLogin()
        {
            Process.Start(new ProcessStartInfo(Executable, "login") { UseShellExecute = true });
        }
    }
}
