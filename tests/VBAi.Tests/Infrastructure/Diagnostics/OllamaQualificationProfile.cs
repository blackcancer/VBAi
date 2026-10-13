using System;
using System.Globalization;
using System.Linq;
using System.Threading.Tasks;

namespace VBAi.Tests.Integration
{
    /// <summary>Explicit test-process profile; never reads or persists personal provider settings.</summary>
    internal sealed class OllamaQualificationProfile
    {
        internal const string TemperatureEnvironmentName = "VBAi_TEST_OLLAMA_TEMPERATURE";
        internal const string TopPEnvironmentName = "VBAi_TEST_OLLAMA_TOP_P";
        internal const string ReasoningEnvironmentName = "VBAi_TEST_OLLAMA_REASONING_EFFORT";
        internal string Model { get; private set; }
        internal Uri Endpoint { get; private set; }
        internal double Temperature { get; private set; }
        internal double TopP { get; private set; }
        internal string ReasoningEffort { get; private set; }

        internal static OllamaQualificationProfile Resolve() => Resolve(Environment.GetEnvironmentVariable);

        internal static OllamaQualificationProfile Resolve(Func<string, string> read)
        {
            if (read == null) throw new ArgumentNullException(nameof(read));
            return new OllamaQualificationProfile
            {
                Model = OllamaQualificationModel.Resolve(read),
                Endpoint = OllamaQualificationEndpoint.Parse(read(OllamaQualificationEndpoint.EnvironmentName)),
                Temperature = ParseSampling(read(TemperatureEnvironmentName), 0, 2, false, 0, TemperatureEnvironmentName),
                TopP = ParseSampling(read(TopPEnvironmentName), 0, 1, true, 0.8, TopPEnvironmentName),
                ReasoningEffort = ParseReasoning(read(ReasoningEnvironmentName))
            };
        }

        private static string ParseReasoning(string configured)
        {
            if (configured == null || configured == OllamaReasoning.Default || configured == OllamaReasoning.None) return configured;
            throw new InvalidOperationException(ReasoningEnvironmentName + " must be default or none; no level is inferred.");
        }

        /// <summary>Creates the production client after checking this model's declarations for an explicit reasoning profile.</summary>
        /// <param name="settings">Disposable settings owned by this synthetic fixture.</param>
        /// <returns>Production HTTP client; no native tool dispatch or personal setting access occurs.</returns>
        internal async Task<LlmChatClient> CreateClientAsync(LlmSettings settings)
        {
            var provider = LlmProvider.All.Single(item => item.IsOllama);
            var capabilities = LlmModelCapabilities.Unknown;
            if (ReasoningEffort != null && ReasoningEffort != OllamaReasoning.Default)
            {
                var models = await LlmChatClient.ListModelsAsync(provider, settings);
                capabilities = models.Single(item => item.Id == Model).Capabilities;
            }
            return new LlmChatClient(provider, settings, Model, null, capabilities, reasoningEffort: ReasoningEffort);
        }

        private static double ParseSampling(string configured, double minimum, double maximum, bool excludeMinimum, double fallback, string name)
        {
            if (configured == null) return fallback;
            double value;
            if (!Double.TryParse(configured.Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out value) ||
                Double.IsNaN(value) || Double.IsInfinity(value) || value < minimum || value > maximum || (excludeMinimum && value == minimum))
                throw new InvalidOperationException(name + " must contain a finite invariant-culture value within the qualification sampling range.");
            return value;
        }

        internal void ApplyTo(LlmSettings settings)
        {
            if (settings == null) throw new ArgumentNullException(nameof(settings));
            settings.OllamaModel = Model;
            settings.OllamaEndpoint = Endpoint.AbsoluteUri;
            settings.OllamaTemperature = Temperature;
            settings.OllamaTopP = TopP;
            if (ReasoningEffort != null) settings.SetReasoningEffort(LlmProvider.All.Single(item => item.IsOllama), Model, ReasoningEffort);
        }

        internal string Describe() => "Model=" + Model + "; Endpoint=" + Endpoint.AbsoluteUri +
            "; Temperature=" + Temperature.ToString("R", CultureInfo.InvariantCulture) +
            "; TopP=" + TopP.ToString("R", CultureInfo.InvariantCulture) + "; Reasoning=" + (ReasoningEffort ?? "provider-default");
    }
}
