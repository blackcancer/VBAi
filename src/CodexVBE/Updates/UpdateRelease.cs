using System;
using System.Linq;
using System.Text.RegularExpressions;

namespace CodexVBE
{
    /// <summary>Version produit indépendante de la version COM, qui reste stable.</summary>
    internal sealed class UpdateVersion : IComparable<UpdateVersion>
    {
        /// <summary>Numeric version components used for precedence comparisons.</summary>
private readonly Version number;
        /// <summary>Prerelease identifiers, or null for a stable release.</summary>
private readonly string preview;
        /// <summary>Gets the original version label without a leading lowercase <c>v</c>.</summary><value>Version text supplied to Parse.</value>
internal string Text { get; }
        /// <summary>Gets whether this version includes a prerelease suffix.</summary><value>True when a prerelease label is present.</value>
internal bool IsPreview => preview != null;
        /// <summary>Creates a parsed version value.</summary><param name="text">Original normalized text.</param><param name="number">Numeric version tuple.</param><param name="preview">Prerelease label, or null for a stable release.</param>
private UpdateVersion(string text, Version number, string preview) { Text = text; this.number = number; this.preview = preview; }
        /// <summary>Parses a semantic version with up to four numeric components and optional prerelease/build labels.</summary>
        /// <param name="text">Version text, optionally prefixed by lowercase <c>v</c>.</param><returns>Parsed version, or null for invalid input.</returns>
internal static UpdateVersion Parse(string text)
        {
            var match = Regex.Match(text ?? "", @"^v?(0|[1-9]\d*)\.(0|[1-9]\d*)\.(0|[1-9]\d*)(?:\.(0|[1-9]\d*))?(?:-([0-9A-Za-z-]+(?:\.[0-9A-Za-z-]+)*))?(?:\+([0-9A-Za-z-]+(?:\.[0-9A-Za-z-]+)*))?$");
            if (!match.Success || text.Length > 128) return null;
            if (!Version.TryParse(string.Join(".", Enumerable.Range(1, 4).Select(i => match.Groups[i].Success ? match.Groups[i].Value : "0")), out var version)) return null;
            string preview = match.Groups[5].Success ? match.Groups[5].Value : null;
            if (preview != null && preview.Split('.').Any(x => x.Length > 1 && x[0] == '0' && x.All(char.IsDigit))) return null;
            return new UpdateVersion(text.TrimStart('v'), version, preview);
        }
        /// <summary>Compares numeric components first, then prerelease identifiers using semantic-version ordering.</summary>
        /// <param name="other">Version to compare.</param><returns>Negative, zero, or positive according to the ordering.</returns>
public int CompareTo(UpdateVersion other)
        {
            if (other == null) return 1;
            int result = number.CompareTo(other.number); if (result != 0) return result;
            if (preview == null || other.preview == null) return preview == other.preview ? 0 : preview == null ? 1 : -1;
            var left = preview.Split('.'); var right = other.preview.Split('.');
            for (int i = 0; i < Math.Min(left.Length, right.Length); i++)
            {
                bool ln = left[i].All(char.IsDigit), rn = right[i].All(char.IsDigit);
                if (ln && rn) { result = left[i].Length.CompareTo(right[i].Length); if (result == 0) result = string.CompareOrdinal(left[i], right[i]); }
                else result = ln == rn ? string.CompareOrdinal(left[i], right[i]) : ln ? -1 : 1;
                if (result != 0) return result;
            }
            return left.Length.CompareTo(right.Length);
        }
    }
    /// <summary>GitHub release metadata used to select and present an updater package.</summary>
internal sealed class UpdateRelease
    {
        /// <summary>Gets or sets the GitHub release tag.</summary><value>Tag parsed as the product version.</value>
public string tag_name { get; set; }
        /// <summary>Gets or sets the release notes body.</summary><value>Markdown notes from the release.</value>
public string body { get; set; }
        /// <summary>Gets or sets whether GitHub marks the release as a draft.</summary><value>Draft state.</value>
public bool draft { get; set; }
        /// <summary>Gets or sets whether GitHub marks the release as a prerelease.</summary><value>Prerelease state.</value>
public bool prerelease { get; set; }
        /// <summary>Gets or sets the files attached to the release.</summary><value>Release assets, or null when none were returned.</value>
public UpdateAsset[] assets { get; set; }
        /// <summary>Parses the release tag into a comparable product version.</summary><value>Parsed version, or null when the tag is invalid.</value>
internal UpdateVersion Version => UpdateVersion.Parse(tag_name);
        /// <summary>Selects the x64 MSI when available, otherwise the x64 setup executable.</summary><value>Supported installer asset, or null when neither exists.</value>
internal UpdateAsset Installer => (assets ?? new UpdateAsset[0]).FirstOrDefault(x => x.name == "VBAi-Setup-win-x64.msi") ??
            (assets ?? new UpdateAsset[0]).FirstOrDefault(x => x.name == "VBAi-Setup-win-x64.exe");
    }
    /// <summary>GitHub file metadata including the digest needed to verify a downloaded package.</summary>
internal sealed class UpdateAsset
    {
        /// <summary>Gets or sets the GitHub asset identifier.</summary><value>Numeric asset ID.</value>
public long id { get; set; }
        /// <summary>Gets or sets the uploaded file name.</summary><value>Asset name.</value>
public string name { get; set; }
        /// <summary>Gets or sets the uploaded file size in bytes.</summary><value>Asset length.</value>
public long size { get; set; }
        /// <summary>Gets or sets the GitHub digest in <c>sha256:</c> format.</summary><value>Digest text, or null if absent.</value>
public string digest { get; set; }
        /// <summary>Extracts a lowercase SHA-256 digest from the GitHub digest field.</summary><value>64-character digest, or null for an unsupported or malformed digest.</value>
internal string Hash => digest != null && Regex.IsMatch(digest, "^sha256:[a-fA-F0-9]{64}$") ? digest.Substring(7).ToLowerInvariant() : null;
    }
}
