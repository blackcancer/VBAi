using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Globalization;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class OllamaQualificationProfileTests
    {
        [TestMethod]
        public void MissingOverridesChooseOnlyTheExplicitTestProfile()
        {
            var profile = OllamaQualificationProfile.Resolve(name => null);
            Assert.AreEqual("qwen2.5:7b-instruct", profile.Model);
            Assert.AreEqual(OllamaQualificationEndpoint.Default, profile.Endpoint.AbsoluteUri);
            Assert.AreEqual(0.0, profile.Temperature); Assert.AreEqual(0.8, profile.TopP);
            var settings = new LlmSettings { ProviderName = "Codex", CodexModel = "preserved", VbeEditApproval = "Never" };
            profile.ApplyTo(settings);
            Assert.AreEqual(profile.Model, settings.OllamaModel); Assert.AreEqual(profile.Endpoint.AbsoluteUri, settings.OllamaEndpoint);
            Assert.AreEqual(0.0, settings.OllamaTemperature); Assert.AreEqual(0.8, settings.OllamaTopP);
            Assert.AreEqual("Codex", settings.ProviderName); Assert.AreEqual("preserved", settings.CodexModel); Assert.AreEqual("Never", settings.VbeEditApproval);
            StringAssert.Contains(profile.Describe(), "Temperature=0"); StringAssert.Contains(profile.Describe(), "TopP=0.8");
        }

        [DataTestMethod, DataRow("0", "0.00001"), DataRow("2", "1"), DataRow(" 0.7 ", " 0.8 "), DataRow("7e-1", "8e-1")]
        public void ExplicitInvariantSamplingAndOwnedEndpointArePreserved(string temperature, string topP)
        {
            var reads = new List<string>();
            var values = new Dictionary<string, string>
            {
                [OllamaQualificationModel.EnvironmentName] = " synthetic/model:tag ",
                [OllamaQualificationEndpoint.EnvironmentName] = "http://127.0.0.1:52541/v1/chat/completions",
                [OllamaQualificationProfile.TemperatureEnvironmentName] = temperature,
                [OllamaQualificationProfile.TopPEnvironmentName] = topP
            };
            var profile = OllamaQualificationProfile.Resolve(name => { reads.Add(name); return values[name]; });
            CollectionAssert.AreEquivalent(new[] { OllamaQualificationModel.EnvironmentName, OllamaQualificationEndpoint.EnvironmentName,
                OllamaQualificationProfile.TemperatureEnvironmentName, OllamaQualificationProfile.TopPEnvironmentName }, reads);
            Assert.AreEqual(4, reads.Count); Assert.AreEqual("synthetic/model:tag", profile.Model); Assert.AreEqual(52541, profile.Endpoint.Port);
            Assert.AreEqual(Double.Parse(temperature, CultureInfo.InvariantCulture), profile.Temperature);
            Assert.AreEqual(Double.Parse(topP, CultureInfo.InvariantCulture), profile.TopP);
        }

        [DataTestMethod, DataRow("", null), DataRow(" ", null), DataRow("NaN", null), DataRow("Infinity", null),
            DataRow("-Infinity", null), DataRow("-0.1", null), DataRow("2.1", null), DataRow("0,7", null),
            DataRow(null, ""), DataRow(null, " "), DataRow(null, "NaN"), DataRow(null, "Infinity"),
            DataRow(null, "-Infinity"), DataRow(null, "0"), DataRow(null, "-0.1"), DataRow(null, "1.1"), DataRow(null, "0,8")]
        public void MalformedNonfiniteAndOutOfRangeOverridesRefuseTheRequestedProfile(string temperature, string topP)
        {
            var error = Assert.ThrowsException<InvalidOperationException>(() => OllamaQualificationProfile.Resolve(name =>
                name == OllamaQualificationProfile.TemperatureEnvironmentName ? temperature :
                name == OllamaQualificationProfile.TopPEnvironmentName ? topP : null));
            StringAssert.Contains(error.Message, temperature != null ? OllamaQualificationProfile.TemperatureEnvironmentName : OllamaQualificationProfile.TopPEnvironmentName);
        }

        [DataTestMethod, DataRow("1", null, 1.0, 0.8), DataRow(null, "0.5", 0.0, 0.5)]
        public void OneSamplingOverrideRetainsTheOtherExplicitTestDefault(string temperature, string topP, double expectedTemperature, double expectedTopP)
        {
            var profile = OllamaQualificationProfile.Resolve(name =>
                name == OllamaQualificationProfile.TemperatureEnvironmentName ? temperature :
                name == OllamaQualificationProfile.TopPEnvironmentName ? topP : null);
            Assert.AreEqual(expectedTemperature, profile.Temperature);
            Assert.AreEqual(expectedTopP, profile.TopP);
        }

        [TestMethod]
        public void ProfileStillRefusesUnsafeEndpointAndMissingDependencies()
        {
            Assert.ThrowsException<InvalidOperationException>(() => OllamaQualificationProfile.Resolve(name =>
                name == OllamaQualificationEndpoint.EnvironmentName ? "http://remote.invalid/v1/chat/completions" : null));
            Assert.ThrowsException<InvalidOperationException>(() => OllamaQualificationProfile.Resolve(name =>
                name == OllamaQualificationModel.EnvironmentName ? " " : null));
            Assert.ThrowsException<ArgumentNullException>(() => OllamaQualificationProfile.Resolve(null));
            Assert.ThrowsException<ArgumentNullException>(() => OllamaQualificationProfile.Resolve(name => null).ApplyTo(null));
            var original = new InvalidOperationException("Synthetic environment reader.");
            Assert.AreSame(original, Assert.ThrowsException<InvalidOperationException>(() => OllamaQualificationProfile.Resolve(name => { throw original; })));
        }

        [TestMethod]
        public void TestProfileSamplingUsesInvariantParsingAndDiagnosticsDespiteCurrentCulture()
        {
            var before = CultureInfo.CurrentCulture;
            try
            {
                CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("fr-FR");
                var profile = OllamaQualificationProfile.Resolve(name => null);
                StringAssert.Contains(profile.Describe(), "TopP=0.8");
                Assert.ThrowsException<InvalidOperationException>(() => OllamaQualificationProfile.Resolve(name =>
                    name == OllamaQualificationProfile.TopPEnvironmentName ? "0,8" : null));
            }
            finally { CultureInfo.CurrentCulture = before; }
        }
    }
}
