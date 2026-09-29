using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace VBAi.Tests.Unit
{
    /// <summary>Vérifie la présentation des options de modèle et d’effort.</summary>
    [TestClass]
    [TestCategory("Unit")]
    public sealed class LlmModelOptionTests
    {
        /// <summary>Préserve l’identité du modèle dans les libellés, valeurs par défaut et efforts.</summary>
        [TestMethod]
        public void LabelsDefaultsAndEffortsPreserveModelIdentity()
        {
            foreach (var label in new[] { null, "", " ", "id" }) {
                var model = new LlmModelOption("id", label);
                Assert.AreEqual("id", model.Id);
                Assert.AreEqual("id", model.Label);
                Assert.AreEqual("id", model.ToString());
                Assert.IsFalse(model.IsDefault);
                Assert.IsNull(model.DefaultEffort);
                Assert.AreEqual(0, model.Efforts.Length);
            }
            var effort = new LlmEffortOption("high", "thorough");
            var efforts = new[] { effort };
            var named = new LlmModelOption("id", "Model", true, "high", efforts);
            Assert.AreEqual("Model (id)", named.ToString());
            Assert.IsTrue(named.IsDefault);
            Assert.AreEqual("high", named.DefaultEffort);
            Assert.AreSame(efforts, named.Efforts);
            Assert.AreEqual("high", effort.Id);
            Assert.AreEqual("thorough", effort.Description);
            Assert.AreEqual("high", effort.ToString());
        }
    }
}
