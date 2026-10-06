// SPDX-License-Identifier: MPL-2.0
using System;
using System.IO;
using System.Text;

namespace VBAi
{
    /// <summary>Attaches the original MIT grant to VBA support distributed into user projects.</summary>
    internal static class VbaRuntimeLicense
    {
        /// <summary>Full MIT notice expressed as CRLF-normalized VBA comments.</summary>
        internal static readonly string Comments = ReadComments();

        /// <summary>Reads the packaged MIT text and makes each line inert VBA commentary.</summary>
        /// <returns>The SPDX marker and complete copyright, permission and warranty notice.</returns>
        /// <exception cref="InvalidOperationException">The assembly was built without its MIT license resource.</exception>
        private static string ReadComments()
        {
            using (var stream = typeof(VbaRuntimeLicense).Assembly.GetManifestResourceStream("VBAi.Licenses.MIT.txt"))
            {
                if (stream == null) throw new InvalidOperationException("The VBA support license resource is missing.");
                using (var reader = new StreamReader(stream, Encoding.UTF8, true))
                {
                    string text = reader.ReadToEnd().Replace("\r\n", "\n").TrimEnd('\n');
                    return "' SPDX-License-Identifier: MIT\r\n' " + text.Replace("\n", "\r\n' ") + "\r\n";
                }
            }
        }
    }
}
