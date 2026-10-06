using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;

namespace VBAi
{

    /// <summary>Describes Git failures using bounded exception and product-frame metadata only.</summary>
    internal static class GitFailureDiagnostic
    {

        /// <summary>Retains error kinds and HRESULTs without copying exception messages or source paths.</summary>
        /// <param name="error">Exception describing the error failure.</param>
        /// <returns>At most four exception type/HRESULT entries with one bounded VBAi frame each; messages and paths are omitted.</returns>
        internal static string Describe(Exception error)
        {
            if (error == null) throw new ArgumentNullException(nameof(error));
            var entries = new List<string>();
            for (Exception current = error; current != null && entries.Count < 4; current = current.InnerException)
            {
                string detail = Bounded(current.GetType().Name, 96) + " 0x" + current.HResult.ToString("X8", CultureInfo.InvariantCulture);
                try
                {
                    var frames = new StackTrace(current, true).GetFrames();
                    if (frames != null)
                        foreach (var frame in frames)
                        {
                            var method = frame.GetMethod();
                            var type = method?.DeclaringType;
                            if (type == null || type.Assembly != typeof(GitFailureDiagnostic).Assembly) continue;
                            detail += " @ " + Bounded(type.FullName + "." + method.Name, 160);
                            int line = frame.GetFileLineNumber();
                            if (line > 0) detail += ":" + line.ToString(CultureInfo.InvariantCulture);
                            break;
                        }
                }
                catch (Exception)
                {
                    // Missing symbols or unavailable frames must not replace the original action failure.
                }
                entries.Add(detail);
            }
            return string.Join(" <- ", entries);
        }

        /// <summary>Bounds metadata independently of messages and runtime object contents.</summary>
        /// <param name="value">Exception type or product method metadata to retain.</param>
        /// <param name="limit">Maximum character count for the emitted metadata field.</param>
        /// <returns>The original string when it fits, otherwise its first <paramref name="limit"/> characters.</returns>
        private static string Bounded(string value, int limit)
        {
            return value.Length <= limit ? value : value.Substring(0, limit);
        }
    }
}
