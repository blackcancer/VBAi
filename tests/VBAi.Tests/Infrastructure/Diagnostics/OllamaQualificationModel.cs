using System;
using System.Linq;

namespace VBAi.Tests.Integration
{
    /// <summary>Selects an already-installed qualification model without loading or saving personal settings.</summary>
    internal static class OllamaQualificationModel
    {
        internal const string EnvironmentName = "VBAi_TEST_OLLAMA_MODEL";
        internal const string Default = "qwen2.5:7b-instruct";

        internal static string Resolve() => Resolve(Environment.GetEnvironmentVariable);

        internal static string Resolve(Func<string, string> read)
        {
            if (read == null) throw new ArgumentNullException(nameof(read));
            return Parse(read(EnvironmentName));
        }

        internal static string Parse(string configured)
        {
            if (configured == null) return Default;
            string model = configured.Trim();
            if (model.Length == 0 || model.Any(character => Char.IsWhiteSpace(character) || Char.IsControl(character)))
                throw new InvalidOperationException("The qualification model override must be nonempty and contain no internal whitespace or control characters.");
            return model;
        }
    }
}
