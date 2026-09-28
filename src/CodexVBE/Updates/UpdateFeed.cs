using System;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Text;

namespace CodexVBE
{
    /// <summary>Lecture publique d’abord ; GCM seulement si GitHub réclame l’accès au dépôt privé.</summary>
    internal sealed class UpdateFeed : IDisposable
    {
        internal const string ApiRoot = "https://api.github.com/repos/blackcancer/CodexVBE";
        internal static Func<LlmSettings> LoadCredentialSettings = LlmSettings.Load;
        internal static Func<string, CancellationToken, Task<string>> ReadCredential = GitHubApi.ReadCredential;
        private static Task<string> DefaultCredential(CancellationToken ct) => ReadCredential(LoadCredentialSettings().GitHubAccount, ct);
        private readonly HttpClient client;
        private readonly Func<CancellationToken, Task<string>> credentials;
        private string token;
        private bool triedCredentials;
        private readonly JavaScriptSerializer json = new JavaScriptSerializer { MaxJsonLength = 4 * 1024 * 1024 };
        internal UpdateFeed(HttpMessageHandler handler = null, Func<CancellationToken, Task<string>> credentials = null)
        {
            client = new HttpClient(handler ?? new HttpClientHandler { AllowAutoRedirect = false }) { Timeout = TimeSpan.FromMinutes(10) };
            this.credentials = credentials ?? DefaultCredential;
        }
        private async Task<HttpResponseMessage> Request(string url, bool binary, CancellationToken ct)
        {
            using (var request = new HttpRequestMessage(HttpMethod.Get, url))
            {
                request.Headers.UserAgent.ParseAdd("VBAi/1.0");
                request.Headers.Accept.ParseAdd(binary ? "application/octet-stream" : "application/vnd.github+json");
                if (new Uri(url).Host == "api.github.com")
                {
                    request.Headers.Add("X-GitHub-Api-Version", "2026-03-10");
                    if (!string.IsNullOrEmpty(token)) request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                }
                return await client.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, ct);
            }
        }
        private async Task<HttpResponseMessage> AuthenticatedRequest(string url, bool binary, CancellationToken ct)
        {
            var response = await Request(url, binary, ct);
            if ((response.StatusCode == HttpStatusCode.NotFound || response.StatusCode == HttpStatusCode.Unauthorized) && !triedCredentials)
            {
                response.Dispose(); triedCredentials = true;
                token = await credentials(ct);
                response = await Request(url, binary, ct);
            }
            return response;
        }
        internal async Task<UpdateRelease> Check(UpdateVersion current, bool previews, string skipped, CancellationToken ct)
        {
            ct.ThrowIfCancellationRequested();
            using (var timeout = CancellationTokenSource.CreateLinkedTokenSource(ct))
            { timeout.CancelAfter(TimeSpan.FromMinutes(2)); return await CheckCore(current, previews, skipped, timeout.Token); }
        }
        internal async Task<string> Download(UpdateAsset asset, string root, IProgress<int> progress, CancellationToken ct)
        {
            ct.ThrowIfCancellationRequested();
            using (var timeout = CancellationTokenSource.CreateLinkedTokenSource(ct))
            { timeout.CancelAfter(TimeSpan.FromMinutes(10)); return await DownloadCore(asset, root, progress, timeout.Token); }
        }
        private static async Task<string> ReadJson(HttpContent content, CancellationToken ct)
        {
            using (var stream = await content.ReadAsStreamAsync())
            using (var bytes = new MemoryStream())
            {
                var buffer = new byte[8192]; int count;
                while ((count = await stream.ReadAsync(buffer, 0, buffer.Length, ct)) > 0)
                { if (bytes.Length + count > 4 * 1024 * 1024) throw new InvalidDataException("Release response is too large."); bytes.Write(buffer, 0, count); }
                return Encoding.UTF8.GetString(bytes.ToArray()).TrimStart('\ufeff');
            }
        }
        private async Task<UpdateRelease> CheckCore(UpdateVersion current, bool previews, string skipped, CancellationToken ct)
        {
            if (current == null) throw new InvalidDataException("Invalid product version.");
            UpdateRelease best = null;
            // Bounded pagination, independent of release date and GitHub's manually selected 'latest'.
            for (int page = 1; page <= 10; page++)
            using (var response = await AuthenticatedRequest(ApiRoot + "/releases?per_page=100&page=" + page, false, ct))
            {
                response.EnsureSuccessStatusCode();
                var releases = json.Deserialize<UpdateRelease[]>(await ReadJson(response.Content, ct)) ?? new UpdateRelease[0];
                ct.ThrowIfCancellationRequested();
                foreach (var release in releases)
                {
                    var version = release.Version;
                    if (release.draft || version == null || (!previews && (release.prerelease || version.IsPreview)) ||
                        version.CompareTo(current) <= 0 || (UpdateVersion.Parse(skipped)?.CompareTo(version) == 0)) continue;
                    if (best == null || version.CompareTo(best.Version) > 0) best = release;
                }
                if (releases.Length < 100) break;
            }
            return best;
        }
        private async Task<string> DownloadCore(UpdateAsset asset, string root, IProgress<int> progress, CancellationToken ct)
        {
            if (asset == null || asset.id <= 0 || asset.Hash == null || asset.size <= 0 || asset.size > 512L * 1024 * 1024)
                throw new InvalidDataException("An installer and its GitHub SHA-256 digest are required.");
            string path = UpdatePaths.AssetPath(root, asset.Hash, asset.name);
            Directory.CreateDirectory(Path.GetDirectoryName(path));
            if (File.Exists(path) && new FileInfo(path).Length == asset.size && UpdatePaths.Hash(path) == asset.Hash) return path;
            string temporary = path + "." + Guid.NewGuid().ToString("N") + ".part";
            try
            {
                string url = ApiRoot + "/releases/assets/" + asset.id;
                for (int redirects = 0; redirects <= 5; redirects++)
                using (var response = redirects == 0 ? await AuthenticatedRequest(url, true, ct) : await Request(url, true, ct))
                {
                    int status = (int)response.StatusCode;
                    if (status >= 300 && status < 400)
                    {
                        var location = response.Headers.Location;
                        var next = location == null ? null : location.IsAbsoluteUri ? location : new Uri(new Uri(url), location);
                        if (next == null || next.Scheme != "https" || !string.IsNullOrEmpty(next.UserInfo) || !next.IsDefaultPort ||
                            (next.Host != "release-assets.githubusercontent.com" && next.Host != "objects.githubusercontent.com"))
                            throw new InvalidDataException("Unexpected GitHub asset redirect.");
                        url = next.AbsoluteUri; continue; // No GitHub token goes to the asset CDN.
                    }
                    response.EnsureSuccessStatusCode();
                    if (response.Content.Headers.ContentLength.HasValue && response.Content.Headers.ContentLength != asset.size) throw new InvalidDataException("Invalid installer size.");
                    using (var source = await response.Content.ReadAsStreamAsync())
                    using (var target = new FileStream(temporary, FileMode.CreateNew, FileAccess.Write, FileShare.None, 81920, true))
                    {
                        var buffer = new byte[81920]; long total = 0; int count;
                        while ((count = await source.ReadAsync(buffer, 0, buffer.Length, ct)) > 0)
                        {
                            total += count; if (total > asset.size) throw new InvalidDataException("Installer exceeds declared size.");
                            await target.WriteAsync(buffer, 0, count, ct); progress?.Report((int)(total * 100 / asset.size));
                        }
                        if (total != asset.size) throw new InvalidDataException("Incomplete installer.");
                    }
                    ct.ThrowIfCancellationRequested();
                    if (UpdatePaths.Hash(temporary) != asset.Hash) throw new InvalidDataException("Installer checksum mismatch.");
                    if (File.Exists(path)) File.Replace(temporary, path, null); else File.Move(temporary, path);
                    return path;
                }
                throw new InvalidDataException("Too many asset redirects.");
            }
            finally { if (File.Exists(temporary)) File.Delete(temporary); }
        }
        public void Dispose() { client.Dispose(); }
    }
}
