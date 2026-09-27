using System;
using System.Diagnostics;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace CodexVBE
{
    internal sealed class GitHubAccountService
    {
        private readonly Func<string, CancellationToken, Task<string>> execute;
        internal GitHubAccountService(Func<string, CancellationToken, Task<string>> execute = null)
        { this.execute = execute ?? Execute; }

        internal async Task<string[]> ListAsync(CancellationToken cancellation)
        {
            return ParseAccounts(await execute("credential-manager github list --url https://github.com --no-ui", cancellation));
        }
        internal async Task LoginAsync(CancellationToken cancellation)
        {
            await execute("credential-manager github login --url https://github.com --browser", cancellation);
        }
        internal static bool ValidAccount(string value)
        {
            return Regex.IsMatch(value ?? "", @"^[A-Za-z0-9](?:[A-Za-z0-9-]{0,37}[A-Za-z0-9])?$");
        }
        internal static string[] ParseAccounts(string text)
        {
            var lines = (text ?? "").Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries).Select(x => x.Trim()).ToArray();
            if (lines.Any(x => !ValidAccount(x)))
                throw new InvalidOperationException(UiText.Get("Unexpected Git Credential Manager response. Check its version."));
            return lines.Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(x => x, StringComparer.OrdinalIgnoreCase).ToArray();
        }
        private static Task<string> Execute(string arguments, CancellationToken cancellation)
        {
            return Task.Run(async () => {
                cancellation.ThrowIfCancellationRequested();
                var start = new ProcessStartInfo("git.exe", arguments) {
                    UseShellExecute = false, CreateNoWindow = true, RedirectStandardOutput = true, RedirectStandardError = true
                };
                using (var process = new Process { StartInfo = start })
                {
                    try { process.Start(); }
                    catch (System.ComponentModel.Win32Exception) { throw new InvalidOperationException(UiText.Get("Git for Windows is required. Install it with Git Credential Manager, then reopen settings.")); }
                    using (cancellation.Register(() => { try { if (!process.HasExited) process.Kill(); } catch (InvalidOperationException) { } catch (System.ComponentModel.Win32Exception) { } }))
                    {
                        var output = process.StandardOutput.ReadToEndAsync();
                        var error = process.StandardError.ReadToEndAsync();
                        if (!process.WaitForExit(300000)) { try { process.Kill(); } catch { } throw new TimeoutException(UiText.Get("GitHub sign-in timed out. Start sign-in again from settings.")); }
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
