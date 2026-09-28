using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Web.Script.Serialization;

namespace CodexVBE
{
    /// <summary>Durable recovery state for the ten native VBE syntax categories.</summary>
    internal sealed class VbeNativePaletteState
    {
        public int Schema { get; set; } = 1;
        public string VbeVersion { get; set; }
        public ColorRow[] Original { get; set; }
        public ColorRow[] Applied { get; set; }

        public sealed class ColorRow
        {
            public string Name { get; set; }
            public int Foreground { get; set; }
            public int Background { get; set; }
            public int Indicator { get; set; }
        }

        internal static ColorRow[] Dark(ColorRow[] original)
        {
            ValidateRows(original);
            int[] foreground = { 8, 16, 13, 1, 16, 3, 4, 8, 8, 1 };
            int[] background = { 1, 2, 1, 15, 5, 1, 1, 1, 1, 15 };
            return original.Select((row, index) => new ColorRow
            {
                Name = row.Name, Foreground = foreground[index], Background = background[index], Indicator = row.Indicator
            }).ToArray();
        }

        internal static void ValidateRows(ColorRow[] rows)
        {
            if (rows == null || rows.Length != 10 || rows.Any(row => row == null || string.IsNullOrWhiteSpace(row.Name) ||
                row.Foreground < 0 || row.Foreground > 16 || row.Background < 0 || row.Background > 16 || row.Indicator < 0 || row.Indicator > 16) ||
                rows.Select(row => row.Name).Distinct(StringComparer.Ordinal).Count() != 10)
                throw new InvalidDataException("The native color palette is incomplete or invalid.");
        }

        internal static bool Equal(ColorRow[] left, ColorRow[] right)
        {
            return left != null && right != null && left.Length == right.Length &&
                left.Zip(right, (a, b) => a != null && b != null && a.Name == b.Name && a.Foreground == b.Foreground &&
                    a.Background == b.Background && a.Indicator == b.Indicator).All(equal => equal);
        }

        internal void Validate(string version)
        {
            if (Schema != 1 || VbeVersion != version) throw new InvalidDataException("The saved native palette belongs to a different VBE version.");
            ValidateRows(Original);
            ValidateRows(Applied);
            if (!Equal(Applied, Dark(Original))) throw new InvalidDataException("The saved native theme palette is inconsistent.");
        }

        internal void RequireUnchanged(ColorRow[] current)
        {
            if (!Equal(current, Original) && !Equal(current, Applied))
                throw new InvalidOperationException("Native editor colors were changed after the theme was applied. The original palette is retained; current colors were not overwritten.");
        }

        internal static VbeNativePaletteState Load(string path, string version)
        {
            if (!File.Exists(path)) return null;
            if (new FileInfo(path).Length > 32768) throw new InvalidDataException("The saved native palette is too large.");
            var state = new JavaScriptSerializer().Deserialize<VbeNativePaletteState>(File.ReadAllText(path));
            if (state == null) throw new InvalidDataException("The saved native palette is empty.");
            state.Validate(version);
            return state;
        }

        internal void SaveNew(string path)
        {
            Validate(VbeVersion);
            Directory.CreateDirectory(Path.GetDirectoryName(path));
            string temporary = path + "." + Guid.NewGuid().ToString("N") + ".tmp";
            try
            {
                byte[] bytes = Encoding.UTF8.GetBytes(new JavaScriptSerializer().Serialize(this));
                using (var file = new FileStream(temporary, FileMode.CreateNew, FileAccess.Write, FileShare.None))
                {
                    file.Write(bytes, 0, bytes.Length);
                    file.Flush(true);
                }
                // Never replace the only copy of the user's original colors.
                File.Move(temporary, path);
            }
            finally { if (File.Exists(temporary)) File.Delete(temporary); }
        }
    }
}
