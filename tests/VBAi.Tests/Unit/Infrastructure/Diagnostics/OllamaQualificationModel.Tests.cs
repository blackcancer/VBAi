using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class OllamaQualificationModelTests
    {
        [TestMethod]
        public void MissingOverrideSelectsTheSharedInstructModel()
        {
            Assert.AreEqual("qwen2.5:7b-instruct", OllamaQualificationModel.Default);
            Assert.AreEqual(OllamaQualificationModel.Default, OllamaQualificationModel.Parse(null));
            Assert.AreEqual(OllamaQualificationModel.Default, OllamaQualificationModel.Resolve(name => null));
        }

        [DataTestMethod, DataRow("qwen3:4b-instruct"), DataRow("qwen2.5:3b"),
            DataRow("synthetic/model:tag"), DataRow("Synthetic.Model:Tag")]
        public void ExplicitModelPreservesNameTagAndCase(string model)
        {
            Assert.AreEqual(model, OllamaQualificationModel.Parse(model));
            Assert.AreEqual(model, OllamaQualificationModel.Resolve(name => model));
        }

        [DataTestMethod, DataRow(" qwen2.5:7b-instruct "), DataRow("\tqwen2.5:7b-instruct\t"),
            DataRow("\r\nqwen2.5:7b-instruct\n"), DataRow("\u00a0qwen2.5:7b-instruct\u2003")]
        public void OuterWhitespaceIsTrimmedWithoutChangingTheModel(string configured)
            => Assert.AreEqual(OllamaQualificationModel.Default, OllamaQualificationModel.Parse(configured));

        [DataTestMethod, DataRow(""), DataRow(" "), DataRow("\t\r\n"), DataRow("\u00a0\u2003")]
        public void EmptyExplicitOverrideIsRefusedRatherThanSilentlySelectingADefault(string configured)
        {
            Assert.ThrowsException<InvalidOperationException>(() => OllamaQualificationModel.Parse(configured));
            Assert.ThrowsException<InvalidOperationException>(() => OllamaQualificationModel.Resolve(name => configured));
        }

        [DataTestMethod, DataRow("qwen3:4b instruct"), DataRow("qwen3:\t4b-instruct"),
            DataRow("qwen3:\n4b-instruct"), DataRow("qwen3:\r4b-instruct"),
            DataRow("qwen3:\u00a04b-instruct"), DataRow("qwen3:\u20034b-instruct"),
            DataRow("qwen3:\\0"), DataRow("\\0qwen3:4b-instruct"),
            DataRow("qwen3:4b-instruct\\0"), DataRow("qwen3:\u007f4b-instruct")]
        public void InternalWhitespaceAndControlCharactersAreRefusedWithoutEchoingConfiguration(string configured)
        {
            // Keep NUL out of the test-case display name and TRX metadata, while testing the actual control character.
            configured = configured.Replace("\\0", "\0");
            var error = Assert.ThrowsException<InvalidOperationException>(() => OllamaQualificationModel.Parse(configured));
            Assert.IsFalse(error.Message.Contains(configured));
            Assert.ThrowsException<InvalidOperationException>(() => OllamaQualificationModel.Resolve(name => configured));
        }

        [TestMethod]
        public void ResolveReadsOnlyTheModelEnvironmentNameExactlyOnce()
        {
            var requested = new List<string>();
            string selected = OllamaQualificationModel.Resolve(name => { requested.Add(name); return " synthetic:installed "; });
            CollectionAssert.AreEqual(new[] { "VBAi_TEST_OLLAMA_MODEL" }, requested);
            Assert.AreEqual("synthetic:installed", selected);
        }

        [TestMethod]
        public void ResolvePropagatesTheOriginalEnvironmentReaderFailure()
        {
            var primary = new InvalidOperationException("Synthetic reader failure.");
            var observed = Assert.ThrowsException<InvalidOperationException>(() => OllamaQualificationModel.Resolve(name => { throw primary; }));
            Assert.AreSame(primary, observed);
        }

        [TestMethod]
        public void MissingReaderIsRefusedBeforeAnyConfigurationIsRead()
            => Assert.ThrowsException<ArgumentNullException>(() => OllamaQualificationModel.Resolve(null));

        [TestMethod]
        public void ParameterlessResolveUsesTheSamePolicyAsTheCurrentProcessOverride()
        {
            string configured = Environment.GetEnvironmentVariable(OllamaQualificationModel.EnvironmentName);
            string expected;
            try { expected = OllamaQualificationModel.Parse(configured); }
            catch (InvalidOperationException)
            {
                Assert.ThrowsException<InvalidOperationException>(() => OllamaQualificationModel.Resolve());
                return;
            }
            Assert.AreEqual(expected, OllamaQualificationModel.Resolve());
        }
    }
}
