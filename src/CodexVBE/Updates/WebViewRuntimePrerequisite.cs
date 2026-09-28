using System;
using System.Diagnostics;
using System.IO;
using System.Net.Http;
using System.Security.Cryptography.X509Certificates;
using System.Threading.Tasks;
using Microsoft.Win32;

namespace CodexVBE
{
    internal sealed class WebViewRuntimePrerequisite
    {
        internal static Func<RegistryHive, RegistryView, IDisposable> OpenRegistryRoot = (hive, view) => RegistryKey.OpenBaseKey(hive, view);
        internal static Func<IDisposable, string, IDisposable> OpenRuntimeKey = (root, path) => ((RegistryKey)root).OpenSubKey(path);
        internal static Func<IDisposable, object> ReadRuntimeVersion = key => ((RegistryKey)key).GetValue("pv");
        internal static Func<HttpClient> CreateClient = NewClient;
        private static HttpClient NewClient() => new HttpClient { Timeout = TimeSpan.FromMinutes(2) };
        internal static Func<string, bool> VerifyTrust = UpdateInstallerRunner.VerifySignatureNative;
        internal static Func<string, X509Certificate2> LoadCertificate = NativeCertificate;
        private static X509Certificate2 NativeCertificate(string path) => new X509Certificate2(X509Certificate.CreateFromSignedFile(path));
        internal static Func<ProcessStartInfo, Process> StartProcess = Process.Start;
        internal const string Bootstrapper = "https://go.microsoft.com/fwlink/p/?LinkId=2124703";
        internal Func<bool> IsInstalled = Installed;
        internal Func<string, Task> Download = DownloadNative;
        internal Func<string, bool> Verify = VerifyMicrosoft;
        internal Func<string, int> Install = InstallNative;
        internal static bool ValidVersion(string value) => Version.TryParse(value, out var version) && version > new Version(0, 0, 0, 0);
        internal static bool Installed()
        {
            foreach (var hive in new[] { RegistryHive.CurrentUser, RegistryHive.LocalMachine })
            foreach (var view in new[] { RegistryView.Registry32, RegistryView.Registry64 })
                using (var root = OpenRegistryRoot(hive, view))
                using (var key = OpenRuntimeKey(root, @"Software\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}"))
                    if (ValidVersion(key == null ? null : ReadRuntimeVersion(key) as string)) return true;
            return false;
        }
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
        private static bool VerifyMicrosoft(string path)
        {
            if (!VerifyTrust(path)) return false;
            using (var certificate = LoadCertificate(path))
                return certificate.GetNameInfo(X509NameType.SimpleName, false) == "Microsoft Corporation";
        }
        private static int InstallNative(string path)
        {
            using (var process = StartProcess(new ProcessStartInfo(path, "/silent /install") { UseShellExecute = false, CreateNoWindow = true }))
            { if (process == null) throw new InvalidOperationException("WebView2 installer did not start."); process.WaitForExit(); return process.ExitCode; }
        }
    }
}
