using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Web.Script.Serialization;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class VbeOptionChangeTests
    {
        [TestMethod]
        public void LegacyFramesRemainSingleEditsAndGroupedFramesRetainTypedUnicodeTargets()
        {
            var json = new JavaScriptSerializer();
            var legacy = json.Deserialize<Request>("{\"Command\":\"set_vbe_option\",\"Pane\":\"Editor\",\"Property\":\"Auto Syntax Check\",\"Value\":false,\"ExpectedOptionsVersion\":\"revision\"}");
            Assert.IsNull(legacy.OptionChanges); Assert.AreEqual(false, legacy.Value);
            var grouped = new Request { Command = legacy.Command, ExpectedOptionsVersion = legacy.ExpectedOptionsVersion,
                OptionChanges = new[] {
                    new VbeOptionChange { Pane = "Éditeur", Property = "Retrait automatique", Value = true },
                    new VbeOptionChange { Pane = "Éditeur", Property = "Largeur de tabulation", Value = 4 },
                    new VbeOptionChange { Pane = "Format de l’éditeur", Property = "Premier plan", Value = "NativeIndex:1", Query = "Texte de commentaire" } } };
            var restored = json.Deserialize<Request>(json.Serialize(grouped));
            Assert.IsNull(restored.Pane); Assert.IsNull(restored.Value);
            Assert.AreEqual(3, restored.OptionChanges.Length);
            Assert.AreEqual("Éditeur", restored.OptionChanges[0].Pane);
            Assert.AreEqual(true, restored.OptionChanges[0].Value);
            Assert.AreEqual(4, restored.OptionChanges[1].Value);
            Assert.AreEqual("NativeIndex:1", restored.OptionChanges[2].Value);
            Assert.AreEqual("Texte de commentaire", restored.OptionChanges[2].Query);
        }
    }
}
