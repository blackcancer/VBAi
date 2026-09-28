using System;
using System.Diagnostics;
using System.IO;
using System.Net.Http;
using System.Security.Cryptography.X509Certificates;
using System.Threading.Tasks;
using Microsoft.Win32;

namespace CodexVBE
{
    /// <summary>Checks, downloads, verifies, and silently installs the Microsoft WebView2 runtime when needed.</summary>
    internal sealed class WebViewRuntimePrerequisite
    {
        /// <summary>Stores the open registry root used by WebViewRuntimePrerequisite.</summary>
        internal static Func<RegistryHive, RegistryView, IDisposable> OpenRegistryRoot = (hive, view) => RegistryKey.OpenBaseKey(hive, view);
        /// <summary>Stores the open runtime key used by WebViewRuntimePrerequisite.</summary>
        internal static Func<IDisposable, string, IDisposable> OpenRuntimeKey = (root, path) => ((RegistryKey)root).OpenSubKey(path);
        /// <summary>Stores the read runtime version used by WebViewRuntimePrerequisite.</summary>
        internal static Func<IDisposable, object> ReadRuntimeVersion = key => ((RegistryKey)key).GetValue("pv");
        /// <summary>Stores the create client used by WebViewRuntimePrerequisite.</summary>
        internal static Func<HttpClient> CreateClient = NewClient;
        /// <summary>Performs the new client operation for WebViewRuntimePrerequisite.</summary>
/// <returns>The result produced by this operation.</returns>
        private static HttpClient NewClient() => new HttpClient { Timeout = TimeSpan.FromMinutes(2) };
        /// <summary>Stores the verify trust used by WebViewRuntimePrerequisite.</summary>
        internal static Func<string, bool> VerifyTrust = UpdateInstallerRunner.VerifySignatureNative;
        /// <summary>Stores the load certificate used by WebViewRuntimePrerequisite.</summary>
        internal static Func<string, X509Certificate2> LoadCertificate = NativeCertificate;
        /// <summary>Performs the native certificate operation for WebViewRuntimePrerequisite.</summary>
/// <param name="path">Text containing the path.</param>
/// <returns>The result produced by this operation.</returns>
        private static X509Certificate2 NativeCertificate(string path) => new X509Certificate2(X509Certificate.CreateFromSignedFile(path));
        /// <summary>Stores the start process used by WebViewRuntimePrerequisite.</summary>
        internal static Func<ProcessStartInfo, Process> StartProcess = Process.Start;
        /// <summary>Microsoft-hosted Evergreen WebView2 bootstrapper URL.</summary>
        internal const string Bootstrapper = "https://go.microsoft.com/fwlink/p/?LinkId=2124703";
        /// <summary>Runtime detection delegate.</summary>
        internal Func<bool> IsInstalled = Installed;
        /// <summary>Bootstrapper download delegate.</summary>
        internal Func<string, Task> Download = DownloadNative;
        /// <summary>Microsoft signature and publisher verification delegate.</summary>
        internal Func<string, bool> Verify = VerifyMicrosoft;
        /// <summary>Silent bootstrapper installation delegate.</summary>
        internal Func<string, int> Install = InstallNative;
        /// <summary>Checks whether a WebView2 runtime version string parses to a positive version.</summary>
        /// <param name="value">Registry version value.</param><returns>Whether it is a valid, nonzero version.</returns>
        internal static bool ValidVersion(string value) => Version.TryParse(value, out var version) && version > new Version(0, 0, 0, 0);
        /// <summary>Checks both registry hives and both registry views for an installed WebView2 runtime.</summary>
        /// <returns><see langword="true"/> when a valid runtime version is registered.</returns>
        internal static bool Installed()
        {
            foreach (var hive in new[] { RegistryHive.CurrentUser, RegistryHive.LocalMachine })
            foreach (var view in new[] { RegistryView.Registry32, RegistryView.Registry64 })
                using (var root = OpenRegistryRoot(hive, view))
                using (var key = OpenRuntimeKey(root, @"Software\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}"))
                    if (ValidVersion(key == null ? null : ReadRuntimeVersion(key) as string)) return true;
            return false;
        }
        /// <summary>Ensures the runtime is installed, verifying the downloaded Microsoft bootstrapper before running it.</summary>
        /// <param name="folder">Parent folder for a temporary bootstrapper directory.</param>
        /// <returns>Task that completes after installation has been verified.</returns>
        /// <exception cref="InvalidDataException">The downloaded file fails signature or publisher verification.</exception>
        /// <exception cref="InvalidOperationException">Installation exits unsuccessfully or the runtime remains absent.</exception>
        internal async Task Ensure(string folder)
        {
            if (IsInstalled()) return;
            string directory = Path.Combine(folder, "WebView2-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(directory);
            string path = Path.Combine(directory, "MicrosoftEdgeWebview2Setup.exe");
            try
            {
                await Download(path);
                if (!Verify(path)) throw new InvalidDataException("WebView2 installer signature is invalid.");
                int code = await Task.Run(() => Install(path));
                if (code != 0 || !IsInstalled()) throw new InvalidOperationException("WebView2 installation could not be verified.");
            }
            finally { if (File.Exists(path)) File.Delete(path); Directory.Delete(directory); }
        }
        /// <summary>Downloads the Microsoft bootstrapper over HTTPS with a two-minute timeout and size limit.</summary>
        /// <param name="path">New file path for the downloaded bootstrapper.</param><returns>Task completed after the file is written.</returns>
        private static async Task DownloadNative(string path)
        {
            using (var client = CreateClient())
            using (var response = await client.GetAsync(Bootstrapper, HttpCompletionOption.ResponseHeadersRead))
            {
                response.EnsureSuccessStatusCode();
                if (response.RequestMessage.RequestUri.Scheme != "https") throw new InvalidDataException("HTTPS required.");
                using (var input = await response.Content.ReadAsStreamAsync())
                using (var output = new FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.None))
                {
                    var buffer = new byte[81920]; int read, size = 0;
                    while ((read = await input.ReadAsync(buffer, 0, buffer.Length)) > 0)
                    { size += read; if (size > 20 * 1024 * 1024) throw new InvalidDataException("Unexpected bootstrapper size."); await output.WriteAsync(buffer, 0, read); }
                }
            }
        }
        /// <summary>Verifies the Authenticode signature and Microsoft Corporation signer name.</summary>
        /// <param name="path">Downloaded bootstrapper path.</param><returns>Whether signature and publisher checks pass.</returns>
        private static bool VerifyMicrosoft(string path)
        {
            if (!VerifyTrust(path)) return false;
            using (var certificate = LoadCertificate(path))
                return certificate.GetNameInfo(X509NameType.SimpleName, false) == "Microsoft Corporation";
        }
        /// <summary>Runs the bootstrapper silently and returns its process exit code.</summary>
        /// <param name="path">Verified bootstrapper path.</param><returns>Process exit code.</returns>
        /// <exception cref="InvalidOperationException">The installer process could not be started.</exception>
        private static int InstallNative(string path)
        {
            using (var process = StartProcess(new ProcessStartInfo(path, "/silent /install") { UseShellExecute = false, CreateNoWindow = true }))
            { if (process == null) throw new InvalidOperationException("WebView2 installer did not start."); process.WaitForExit(); return process.ExitCode; }
        }
    }
}
