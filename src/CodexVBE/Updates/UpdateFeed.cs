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
        /// <summary>GitHub API endpoint for the product repository.</summary>
internal const string ApiRoot = "https://api.github.com/repos/blackcancer/CodexVBE";
        /// <summary>Stores the load credential settings used by UpdateFeed.</summary>
internal static Func<LlmSettings> LoadCredentialSettings = LlmSettings.Load;
        /// <summary>Stores the read credential used by UpdateFeed.</summary>
internal static Func<string, CancellationToken, Task<string>> ReadCredential = GitHubApi.ReadCredential;
        /// <summary>Performs the default credential operation for UpdateFeed.</summary>
/// <param name="ct">Token used to cancel the operation.</param>
/// <returns>The result produced by this operation.</returns>
private static Task<string> DefaultCredential(CancellationToken ct) => ReadCredential(LoadCredentialSettings().GitHubAccount, ct);
        /// <summary>HTTP client used for release metadata and asset downloads.</summary>
private readonly HttpClient client;
        /// <summary>Optional credential provider invoked after an unauthenticated private-repository response.</summary>
private readonly Func<CancellationToken, Task<string>> credentials;
        /// <summary>GitHub token cached for authenticated API requests, or null.</summary>
private string token;
        /// <summary>Whether this feed has already attempted credential acquisition.</summary>
private bool triedCredentials;
        /// <summary>JSON serializer with the accepted release-response size limit.</summary>
private readonly JavaScriptSerializer json = new JavaScriptSerializer { MaxJsonLength = 4 * 1024 * 1024 };
        /// <summary>Creates a feed with optional HTTP and credential providers.</summary>
        /// <param name="handler">Optional HTTP handler for requests.</param><param name="credentials">Optional token provider used for private access.</param>
internal UpdateFeed(HttpMessageHandler handler = null, Func<CancellationToken, Task<string>> credentials = null)
        {
            client = new HttpClient(handler ?? new HttpClientHandler { AllowAutoRedirect = false }) { Timeout = TimeSpan.FromMinutes(10) };
            this.credentials = credentials ?? DefaultCredential;
        }
        /// <summary>Sends one unauthenticated or currently-token-authenticated GET request.</summary>
        /// <param name="url">Absolute request URL.</param><param name="binary">Whether to request an installer asset.</param><param name="ct">Cancellation token.</param>
        /// <returns>HTTP response owned by the caller.</returns>
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
        /// <summary>Retries one 404 or 401 response once after obtaining repository credentials.</summary>
        /// <param name="url">Absolute request URL.</param><param name="binary">Whether to request binary content.</param><param name="ct">Cancellation token.</param>
        /// <returns>Final HTTP response owned by the caller.</returns>
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
        /// <summary>Checks GitHub releases with a two-minute timeout and returns the newest eligible version.</summary>
        /// <param name="current">Installed version used as the lower bound.</param><param name="previews">Whether prereleases are eligible.</param>
        /// <param name="skipped">Version the user chose to skip, or null.</param><param name="ct">Cancellation token.</param>
        /// <returns>Newest eligible release, or null when none is newer.</returns>
internal async Task<UpdateRelease> Check(UpdateVersion current, bool previews, string skipped, CancellationToken ct)
        {
            ct.ThrowIfCancellationRequested();
            using (var timeout = CancellationTokenSource.CreateLinkedTokenSource(ct))
            { timeout.CancelAfter(TimeSpan.FromMinutes(2)); return await CheckCore(current, previews, skipped, timeout.Token); }
        }
        /// <summary>Downloads one validated installer with a ten-minute timeout and verifies its length and digest.</summary>
        /// <param name="asset">GitHub asset metadata and expected SHA-256 digest.</param><param name="root">Update cache root.</param>
        /// <param name="progress">Optional percentage progress reporter.</param><param name="ct">Cancellation token.</param>
        /// <returns>Path to the verified cached package.</returns>
internal async Task<string> Download(UpdateAsset asset, string root, IProgress<int> progress, CancellationToken ct)
        {
            ct.ThrowIfCancellationRequested();
            using (var timeout = CancellationTokenSource.CreateLinkedTokenSource(ct))
            { timeout.CancelAfter(TimeSpan.FromMinutes(10)); return await DownloadCore(asset, root, progress, timeout.Token); }
        }
        /// <summary>Reads a bounded UTF-8 JSON response and strips an optional byte-order mark.</summary>
        /// <param name="content">HTTP response content.</param><param name="ct">Cancellation token.</param><returns>Decoded JSON text.</returns>
        /// <exception cref="InvalidDataException">The response exceeds the 4 MiB limit.</exception>
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
        /// <summary>Pages through a bounded list of GitHub releases and selects the highest eligible version.</summary>
        /// <param name="current">Installed version lower bound.</param><param name="previews">Whether prereleases are eligible.</param>
        /// <param name="skipped">Version the user chose to skip, or null.</param><param name="ct">Cancellation token.</param>
        /// <returns>Newest eligible release, or null.</returns>
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
        /// <summary>Downloads a bounded installer through validated HTTPS redirects and verifies its SHA-256 digest.</summary>
        /// <param name="asset">Expected GitHub asset metadata.</param><param name="root">Update cache root.</param>
        /// <param name="progress">Optional percentage progress reporter.</param><param name="ct">Cancellation token.</param>
        /// <returns>Path to the verified cached package.</returns>
        /// <exception cref="InvalidDataException">Asset metadata, redirect, size, or digest validation fails.</exception>
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
        /// <summary>Disposes the HTTP client owned by this feed.</summary>
public void Dispose() { client.Dispose(); }
    }
}
