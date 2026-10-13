using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Linq;

namespace VBAi.Tests.Unit
{
    /// <summary>Exercises explicit default/off controls without guessed model-specific effort levels.</summary>
    [TestClass, TestCategory("Unit")]
    public sealed class OllamaReasoningTests
    {
        [TestMethod]
        public void OptionsRequirePositiveDeclarationsAndNeverInventNamedLevels()
        {
            foreach (var capabilities in new[] { null, LlmModelCapabilities.Unknown, new LlmModelCapabilities(reasoning: false) })
            {
                Assert.AreEqual(0, OllamaReasoning.Options(capabilities).Length);
                Assert.IsNull(OllamaReasoning.Validate(null, capabilities));
                Assert.IsNull(OllamaReasoning.Validate("default", capabilities));
                Assert.ThrowsException<InvalidOperationException>(() => OllamaReasoning.Validate("none", capabilities));
            }
            var supported = new LlmModelCapabilities(reasoning: true);
            var options = OllamaReasoning.Options(supported);
            CollectionAssert.AreEqual(new[] { "default", "none" }, options.Select(item => item.Id).ToArray());
            Assert.AreEqual("none", OllamaReasoning.Validate("none", supported));
            Assert.IsNull(OllamaReasoning.Validate("default", supported));
            Assert.IsTrue(options.All(item => !string.IsNullOrWhiteSpace(item.Description)));
            Assert.AreNotEqual(options[0].Id, options[0].ToString());
            Assert.AreNotEqual(options[1].Id, options[1].ToString());
        }

        [DataTestMethod, DataRow(""), DataRow(" "), DataRow("None"), DataRow("low"), DataRow("medium"), DataRow("high"), DataRow("max"), DataRow("injected")]
        public void UndeclaredSelectionsAreRejectedRatherThanCoercedToDefaults(string selected)
        {
            Assert.ThrowsException<InvalidOperationException>(() => OllamaReasoning.Validate(selected, new LlmModelCapabilities(reasoning: true)));
        }
    }
}
