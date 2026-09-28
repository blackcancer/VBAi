using System;
using System.Linq;
using System.Text.RegularExpressions;

namespace CodexVBE
{
    /// <summary>Version produit indépendante de la version COM, qui reste stable.</summary>
    internal sealed class UpdateVersion : IComparable<UpdateVersion>
    {
        private readonly Version number;
        private readonly string preview;
        internal string Text { get; }
        internal bool IsPreview => preview != null;
        private UpdateVersion(string text, Version number, string preview) { Text = text; this.number = number; this.preview = preview; }
        internal static UpdateVersion Parse(string text)
        {
            var match = Regex.Match(text ?? "", @"^v?(0|[1-9]\d*)\.(0|[1-9]\d*)\.(0|[1-9]\d*)(?:\.(0|[1-9]\d*))?(?:-([0-9A-Za-z-]+(?:\.[0-9A-Za-z-]+)*))?(?:\+([0-9A-Za-z-]+(?:\.[0-9A-Za-z-]+)*))?$");
            if (!match.Success || text.Length > 128) return null;
            if (!Version.TryParse(string.Join(".", Enumerable.Range(1, 4).Select(i => match.Groups[i].Success ? match.Groups[i].Value : "0")), out var version)) return null;
            string preview = match.Groups[5].Success ? match.Groups[5].Value : null;
            if (preview != null && preview.Split('.').Any(x => x.Length > 1 && x[0] == '0' && x.All(char.IsDigit))) return null;
            return new UpdateVersion(text.TrimStart('v'), version, preview);
        }
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
    internal sealed class UpdateRelease
    {
        public string tag_name { get; set; }
        public string body { get; set; }
        public bool draft { get; set; }
        public bool prerelease { get; set; }
        public UpdateAsset[] assets { get; set; }
        internal UpdateVersion Version => UpdateVersion.Parse(tag_name);
        internal UpdateAsset Installer => (assets ?? new UpdateAsset[0]).FirstOrDefault(x => x.name == "VBAi-Setup-win-x64.msi") ??
            (assets ?? new UpdateAsset[0]).FirstOrDefault(x => x.name == "VBAi-Setup-win-x64.exe");
    }
    internal sealed class UpdateAsset
    {
        public long id { get; set; }
        public string name { get; set; }
        public long size { get; set; }
        public string digest { get; set; }
        internal string Hash => digest != null && Regex.IsMatch(digest, "^sha256:[a-fA-F0-9]{64}$") ? digest.Substring(7).ToLowerInvariant() : null;
    }
}
