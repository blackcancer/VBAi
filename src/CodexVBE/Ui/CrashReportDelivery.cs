using System;
using System.Diagnostics;
using System.Net.Http;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Win32;

namespace CodexVBE
{
    /// <summary>Résultat explicite : une création incertaine ne doit jamais déclencher un second envoi automatique.</summary>
    internal enum CrashDeliveryResult { GitHub, Outlook, Draft, Uncertain }

    /// <summary>Publie sur le dépôt du produit avec GCM, ou utilise Outlook puis le client mail local.</summary>
    internal sealed class CrashReportDelivery
    {
        internal static Func<LlmSettings> LoadSettings = LlmSettings.Load;
        internal static Func<string, CancellationToken, Task<string>> ReadCredential = GitHubApi.ReadCredential;
        internal static Func<string, Func<CancellationToken, Task<string>>, GitHubApi> CreateApi =
            (account, credential) => new GitHubApi(account, credential: credential);
        internal static Func<string, object> ActiveOutlook = Marshal.GetActiveObject;
        internal static Func<string, Type> OutlookType = Type.GetTypeFromProgID;
        internal static Func<Type, object> CreateOutlook = Activator.CreateInstance;
        internal static Func<object, bool> IsComReference = Marshal.IsComObject;
        internal static Func<object, int> ReleaseReference = Marshal.ReleaseComObject;
        internal static Func<string, IDisposable> OpenProfiles = path => Registry.CurrentUser.OpenSubKey(path);
        internal static Func<IDisposable, int> ReadProfileSubKeys = profiles => ((RegistryKey)profiles).SubKeyCount;
        internal static Func<string, int?> ProfileCount = version =>
        {
            using (var profiles = OpenProfiles(@"Software\Microsoft\Office" + version + @"\Outlook\Profiles"))
                return profiles == null ? (int?)null : ReadProfileSubKeys(profiles);
        };
        internal Func<string, string, CancellationToken, Task<string>> Publish = PublishNative;
        internal Func<string, string, bool> SendOutlook = SendOutlookNative;
        internal Action<string> OpenDraft = url => Process.Start(new ProcessStartInfo(url) { UseShellExecute = true });
        internal string IssueUrl { get; private set; }

        internal async Task<CrashDeliveryResult> Send(string title, string body, string savedPath, CancellationToken ct)
        {
            try
            {
                IssueUrl = await Publish(title, body, ct);
                if (!Uri.TryCreate(IssueUrl, UriKind.Absolute, out var uri) || uri.Scheme != "https" ||
                    uri.Host != "github.com" || !uri.AbsolutePath.StartsWith("/blackcancer/CodexVBE/issues/", StringComparison.OrdinalIgnoreCase))
                    return CrashDeliveryResult.Uncertain;
                return CrashDeliveryResult.GitHub;
            }
            catch (GitHubApiFailure error) when (error.Status >= 400 && error.Status < 500) { }
            catch (CrashCredentialUnavailable) { }
            catch (OperationCanceledException) { return CrashDeliveryResult.Uncertain; }
            catch (Exception) { return CrashDeliveryResult.Uncertain; }
            ct.ThrowIfCancellationRequested();
            try { return Email(title, body, savedPath); }
            catch (CrashMailUncertain) { return CrashDeliveryResult.Uncertain; }
        }

        internal CrashDeliveryResult Email(string title, string body, string savedPath)
        {
            if (SendOutlook(title, body)) return CrashDeliveryResult.Outlook;
            // Keep mailto short enough for Windows clients. The complete report remains in the file and clipboard.
            string note = UiText.Get("Attach the saved report to this email.") + "\r\n\r\n" + savedPath;
            OpenDraft("mailto:" + CrashReport.Recipient + "?subject=" + Uri.EscapeDataString(title) + "&body=" + Uri.EscapeDataString(note));
            return CrashDeliveryResult.Draft;
        }

        private static async Task<string> PublishNative(string title, string body, CancellationToken ct)
        {
            string account;
            try { account = LoadSettings().GitHubAccount; }
            catch (Exception) { throw new CrashCredentialUnavailable(); }
            using (var api = CreateApi(account, async token =>
            {
                try { return await ReadCredential(account, token); }
                catch (OperationCanceledException) when (!token.IsCancellationRequested) { throw new CrashCredentialUnavailable(); }
                catch (OperationCanceledException) { throw; }
                catch (Exception) { throw new CrashCredentialUnavailable(); }
            }))
            {
                var result = await api.CreateIssue(CrashReport.Repository, title, body, ct);
                return result?.html_url;
            }
        }

        private static bool HasOutlookProfile()
        {
            // Avoid starting Outlook's account setup wizard while recovering a crash.
            foreach (string version in new[] { "16.0", "15.0", "14.0" })
                if ((ProfileCount(version) ?? 0) > 0) return true;
            return false;
        }

        private static bool SendOutlookNative(string title, string body)
        {
            object application = null, session = null, accounts = null, mail = null;
            bool sending = false;
            try
            {
                try { application = ActiveOutlook("Outlook.Application"); }
                catch (COMException)
                {
                    if (!HasOutlookProfile()) return false;
                    var type = OutlookType("Outlook.Application");
                    if (type == null) return false;
                    application = CreateOutlook(type);
                }
                session = ((dynamic)application).Session;
                accounts = ((dynamic)session).Accounts;
                if ((int)((dynamic)accounts).Count == 0) return false;
                mail = ((dynamic)application).CreateItem(0);
                ((dynamic)mail).To = CrashReport.Recipient;
                ((dynamic)mail).Subject = title;
                ((dynamic)mail).Body = body;
                sending = true;
                ((dynamic)mail).Send();
                return true;
            }
            catch (Exception)
            {
                // A failed Send may have queued the message. Do not silently create a duplicate draft.
                if (sending) throw new CrashMailUncertain();
                return false;
            }
            finally
            {
                foreach (var item in new[] { mail, accounts, session, application })
                    try { if (item != null && IsComReference(item)) ReleaseReference(item); }
                    catch (Exception) { LoadLog.Write("Outlook report COM reference could not be released."); }
            }
        }
    }

    /// <summary>La demande n’a pas été envoyée car aucun identifiant utilisable n’a été obtenu.</summary>
    internal sealed class CrashCredentialUnavailable : Exception { }
    /// <summary>Outlook a été sollicité mais son état final ne peut être établi.</summary>
    internal sealed class CrashMailUncertain : Exception { }
}
