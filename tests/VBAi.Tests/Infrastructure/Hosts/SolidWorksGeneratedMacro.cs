using System;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;

namespace VBAi.Tests.Infrastructure
{
    /// <summary>Plans removal of only the known SOLIDWORKS-generated default procedure from a newly owned test macro.</summary>
    /// <remarks>The caller must first prove fresh native creation, exact host/candidate identity and
    /// retain the complete original source. This helper never opens, saves or executes a macro.
    /// Existing qualification fixtures still require blank source and must not invoke this planner.</remarks>
    internal static class SolidWorksGeneratedMacro
    {
        /// <summary>Builds one hash-guarded edit for a proven fresh standard module, or no edit for blank source.</summary>
        /// <param name="path">Absolute path of the newly created disposable SWP.</param>
        /// <param name="module">Exact standard-module name from the owned project inventory.</param>
        /// <param name="type">VBIDE component type; only a standard module is accepted.</param>
        /// <param name="mode">Observed project mode; must be design mode.</param>
        /// <param name="source">Complete source retained before planning the edit.</param>
        /// <param name="sha256">Source checksum returned by the current candidate.</param>
        /// <param name="lineCount">Current native code-module line count.</param>
        /// <returns>A single replace request, or null if the module is already blank.</returns>
        /// <exception cref="InvalidOperationException">The source or metadata is not the known fresh fixture.</exception>
        internal static Request PlanBlankModule(string path, string module, int type, int mode,
            string source, string sha256, int lineCount)
        {
            if (string.IsNullOrWhiteSpace(path) || !Path.IsPathRooted(path) ||
                !string.Equals(Path.GetExtension(path), ".swp", StringComparison.OrdinalIgnoreCase) ||
                string.IsNullOrWhiteSpace(module) || type != 1 || mode != 2 ||
                source == null || lineCount < 0 || lineCount > 128)
                throw new InvalidOperationException("Only an exact fresh standard module in a disposable design-mode SWP may be prepared.");
            string actualHash;
            using (var hash = SHA256.Create())
                actualHash = BitConverter.ToString(hash.ComputeHash(Encoding.UTF8.GetBytes(source))).Replace("-", "");
            if (!string.Equals(actualHash, sha256, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("Generated source checksum differs; no fixture edit is admitted.");
            string[] lines = source.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n')
                .Select(line => line.Trim()).Where(line => line.Length > 0).ToArray();
            if (lines.Length == 0 || (lines.Length == 1 && lines[0].Equals("Option Explicit", StringComparison.OrdinalIgnoreCase)))
                return null;
            string[] generated = { "Dim swApp As Object", "Sub main()", "Set swApp = Application.SldWorks", "End Sub" };
            if (!lines.SequenceEqual(generated, StringComparer.OrdinalIgnoreCase) ||
                lineCount < source.TrimEnd('\r', '\n').Split('\n').Length)
                throw new InvalidOperationException("The module is not the known SOLIDWORKS-generated default; no arbitrary code is removed.");
            return new Request
            {
                Command = "replace_lines", Project = Path.GetFullPath(path), Module = module,
                ExpectedMode = 2, ExpectedSha256 = sha256, StartLine = 1, Count = lineCount,
                Text = "Option Explicit\r\n"
            };
        }
    }
}
