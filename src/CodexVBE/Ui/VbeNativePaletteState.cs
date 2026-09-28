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
        /// <summary>Gets or sets the serialized recovery schema version.</summary>
        /// <value>Schema version understood by this add-in.</value>
        public int Schema { get; set; } = 1;
        /// <summary>Gets or sets the VBE version whose palette was captured.</summary>
        /// <value>Version string used to scope the saved colors.</value>
        public string VbeVersion { get; set; }
        /// <summary>Gets or sets the original ten syntax-category colors.</summary>
        /// <value>Original palette rows in VBE category order.</value>
        public ColorRow[] Original { get; set; }
        /// <summary>Gets or sets the dark palette that was applied.</summary>
        /// <value>Configured dark palette rows in VBE category order.</value>
        public ColorRow[] Applied { get; set; }

        /// <summary>Color-index values and category name captured from one VBE syntax row.</summary>
        public sealed class ColorRow
        {
            /// <summary>Gets or sets the VBE syntax-category name.</summary>
            /// <value>Unique category label.</value>
            public string Name { get; set; }
            /// <summary>Gets or sets the foreground palette index.</summary>
            /// <value>Index of the foreground color in the VBE palette.</value>
            public int Foreground { get; set; }
            /// <summary>Gets or sets the background palette index.</summary>
            /// <value>Index of the background color in the VBE palette.</value>
            public int Background { get; set; }
            /// <summary>Gets or sets the indicator palette index.</summary>
            /// <value>Index of the indicator color in the VBE palette.</value>
            public int Indicator { get; set; }
        }

        /// <summary>Builds the configured dark palette while retaining each category's indicator color.</summary>
        /// <param name="original">The validated original palette, in VBE category order.</param>
        /// <returns>A new array containing the dark foreground and background indices.</returns>
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

        /// <summary>Checks that the palette contains ten uniquely named rows with valid color indices.</summary>
        /// <param name="rows">The palette rows to validate.</param>
        /// <exception cref="InvalidDataException">The rows are missing, duplicated, or contain an invalid index.</exception>
        internal static void ValidateRows(ColorRow[] rows)
        {
            if (rows == null || rows.Length != 10 || rows.Any(row => row == null || string.IsNullOrWhiteSpace(row.Name) ||
                row.Foreground < 0 || row.Foreground > 16 || row.Background < 0 || row.Background > 16 || row.Indicator < 0 || row.Indicator > 16) ||
                rows.Select(row => row.Name).Distinct(StringComparer.Ordinal).Count() != 10)
                throw new InvalidDataException("The native color palette is incomplete or invalid.");
        }

        /// <summary>Compares two palettes row by row, including names and all three color indices.</summary>
        /// <param name="left">The first palette.</param>
        /// <param name="right">The second palette.</param>
        /// <returns><see langword="true"/> when both arrays have equal rows in the same order.</returns>
        internal static bool Equal(ColorRow[] left, ColorRow[] right)
        {
            return left != null && right != null && left.Length == right.Length &&
                left.Zip(right, (a, b) => a != null && b != null && a.Name == b.Name && a.Foreground == b.Foreground &&
                    a.Background == b.Background && a.Indicator == b.Indicator).All(equal => equal);
        }

        /// <summary>Validates the schema, VBE version, rows, and dark-palette derivation.</summary>
        /// <param name="version">The VBE version expected by the caller.</param>
        /// <exception cref="InvalidDataException">The saved state is incompatible or internally inconsistent.</exception>
        internal void Validate(string version)
        {
            if (Schema != 1 || VbeVersion != version) throw new InvalidDataException("The saved native palette belongs to a different VBE version.");
            ValidateRows(Original);
            ValidateRows(Applied);
            if (!Equal(Applied, Dark(Original))) throw new InvalidDataException("The saved native theme palette is inconsistent.");
        }

        /// <summary>Refuses recovery if current colors match neither the saved original nor applied palette.</summary>
        /// <param name="current">The palette read from the VBE before recovery.</param>
        /// <exception cref="InvalidOperationException">The user changed editor colors after the theme was applied.</exception>
        internal void RequireUnchanged(ColorRow[] current)
        {
            if (!Equal(current, Original) && !Equal(current, Applied))
                throw new InvalidOperationException("Native editor colors were changed after the theme was applied. The original palette is retained; current colors were not overwritten.");
        }

        /// <summary>Loads and validates a saved palette state, returning <see langword="null"/> when absent.</summary>
        /// <param name="path">Path to the recovery file.</param>
        /// <param name="version">VBE version that must match the saved state.</param>
        /// <returns>The validated state, or <see langword="null"/> if no file exists.</returns>
        /// <exception cref="InvalidDataException">The file is oversized, empty, or invalid.</exception>
        internal static VbeNativePaletteState Load(string path, string version)
        {
            if (!File.Exists(path)) return null;
            if (new FileInfo(path).Length > 32768) throw new InvalidDataException("The saved native palette is too large.");
            var state = new JavaScriptSerializer().Deserialize<VbeNativePaletteState>(File.ReadAllText(path));
            if (state == null) throw new InvalidDataException("The saved native palette is empty.");
            state.Validate(version);
            return state;
        }

        /// <summary>Creates a new recovery file without replacing an existing copy of the original palette.</summary>
        /// <param name="path">Destination path for the serialized state.</param>
        /// <exception cref="InvalidDataException">The state fails validation before writing.</exception>
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
