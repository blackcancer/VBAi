using Microsoft.Win32;
using System;
using System.Diagnostics;
using System.IO;
using System.Net.Http;
using System.Security.Cryptography.X509Certificates;
using System.Threading;
using System.Threading.Tasks;

namespace VBAi
{

    /// <summary>Checks, downloads, verifies, and silently installs the Microsoft WebView2 runtime when needed.</summary>
    internal sealed class WebViewRuntimePrerequisite
    {

        /// <summary>Opens one requested registry hive/view for detection; the returned handle is disposed by Installed.</summary>
        internal static Func<RegistryHive, RegistryView, IDisposable> OpenRegistryRoot = (hive, view) => RegistryKey.OpenBaseKey(hive, view);

        /// <summary>Opens the WebView2 client subkey read-only below a detection root, returning null when it is absent.</summary>
        internal static Func<IDisposable, string, IDisposable> OpenRuntimeKey = (root, path) => ((RegistryKey)root).OpenSubKey(path);

        /// <summary>Reads the pv registry value from the detected runtime client key; non-string values do not establish installation.</summary>
        internal static Func<IDisposable, object> ReadRuntimeVersion = key => ((RegistryKey)key).GetValue("pv");

        /// <summary>Creates the owned HTTP client for one bootstrapper download; the download disposes it.</summary>
        internal static Func<HttpClient> CreateClient = NewClient;

        /// <summary>Creates a bootstrapper HTTP client with a two-minute request timeout.</summary>
        /// <returns>A new client that the download routine owns and disposes.</returns>
        private static HttpClient NewClient() => new HttpClient { Timeout = TimeSpan.FromMinutes(2) };

        /// <summary>Verifies bootstrapper Authenticode trust before its Microsoft signer name is accepted.</summary>
        internal static Func<string, bool> VerifyTrust = UpdateInstallerRunner.VerifySignatureNative;

        /// <summary>Loads the bootstrapper signer certificate; VerifyMicrosoft disposes the returned certificate.</summary>
        internal static Func<string, X509Certificate2> LoadCertificate = NativeCertificate;

        /// <summary>Loads the signer's certificate from a signed bootstrapper file without independently verifying trust.</summary>
        /// <param name="path">Path of the downloaded signed bootstrapper whose certificate is requested.</param>
        /// <returns>A caller-owned signer certificate; malformed or unsigned files raise the certificate API error.</returns>
        private static X509Certificate2 NativeCertificate(string path) => new X509Certificate2(X509Certificate.CreateFromSignedFile(path));

        /// <summary>Starts the verified bootstrapper using its explicit silent arguments; the returned process handle is disposed by InstallNative.</summary>
        internal static Func<ProcessStartInfo, Process> StartProcess = Process.Start;

        /// <summary>Bounds the silent installer wait without terminating a process with an uncertain outcome.</summary>
        internal static Func<Process, int, bool> WaitForInstaller = (process, milliseconds) => process.WaitForExit(milliseconds);

        /// <summary>Maximum silent-installer wait, fifteen minutes expressed in milliseconds; expiration does not terminate or retry the installer.</summary>
        internal const int InstallerTimeoutMilliseconds = 15 * 60 * 1000;

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
        /// <exception cref="TimeoutException">The started installer did not exit within its wait; its files are retained and it is not terminated or retried.</exception>
        internal async Task Ensure(string folder)
        {
            if (IsInstalled()) return;
            string directory = Path.Combine(folder, "WebView2-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(directory);
            string path = Path.Combine(directory, "MicrosoftEdgeWebview2Setup.exe");
            Exception failure = null;
            bool installerTimedOut = false;
            try
            {
                await Download(path);
                if (!Verify(path)) throw new InvalidDataException("WebView2 installer signature is invalid.");
                int code;
                try { code = await Task.Run(() => Install(path)); }
                catch (TimeoutException) { installerTimedOut = true; throw; }
                if (code != 0 || !IsInstalled()) throw new InvalidOperationException("WebView2 installation could not be verified.");
            }
            catch (Exception error) { failure = error; throw; }
            finally
            {
                if (!installerTimedOut)
                {
                    try { if (File.Exists(path)) File.Delete(path); Directory.Delete(directory); }
                    catch (IOException) when (failure != null) { }
                    catch (UnauthorizedAccessException) when (failure != null) { }
                }
            }
        }

        /// <summary>Downloads the Microsoft bootstrapper over HTTPS with a two-minute timeout and size limit.</summary>
        /// <param name="path">New file path for the downloaded bootstrapper.</param><returns>Task completed after the file is written.</returns>
        private static async Task DownloadNative(string path)
        {
            using (var timeout = new CancellationTokenSource(TimeSpan.FromMinutes(2)))
            using (var client = CreateClient())
            using (var response = await client.GetAsync(Bootstrapper, HttpCompletionOption.ResponseHeadersRead, timeout.Token))
            {
                response.EnsureSuccessStatusCode();
                if (response.RequestMessage.RequestUri.Scheme != "https") throw new InvalidDataException("HTTPS required.");
                using (var input = await response.Content.ReadAsStreamAsync())
                using (var output = new FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.None))
                {
                    var buffer = new byte[81920]; int read, size = 0;
                    while ((read = await input.ReadAsync(buffer, 0, buffer.Length, timeout.Token)) > 0)
                    { size += read; if (size > 20 * 1024 * 1024) throw new InvalidDataException("Unexpected bootstrapper size."); await output.WriteAsync(buffer, 0, read, timeout.Token); }
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
        /// <exception cref="TimeoutException">The original process handle did not signal exit before the deadline; the child is left running.</exception>
        private static int InstallNative(string path)
        {
            using (var process = StartProcess(new ProcessStartInfo(path, "/silent /install") { UseShellExecute = false, CreateNoWindow = true }))
            {
                if (process == null) throw new InvalidOperationException("WebView2 installer did not start.");
                if (!WaitForInstaller(process, InstallerTimeoutMilliseconds))
                    throw new TimeoutException("WebView2 installation outcome is uncertain. Check the installer and runtime state before retrying.");
                return process.ExitCode;
            }
        }
    }
}
