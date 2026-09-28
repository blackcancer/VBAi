using CodexVBE;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace CodexVBE.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class UpdateTextTests
    {
        [TestMethod]
        public void InstallerTextUsesOwnedCultureAndPreservesUnknownKeys()
        {
            string original = UpdateText.Culture;
            try
            {
                foreach (string culture in new[] { "en-US", "fr-FR", "ar-SA" })
                {
                    UpdateText.Culture = culture;
                    Assert.AreEqual("Owned missing update key", UpdateText.Get("Owned missing update key"));
                    string translated = UpdateText.Get("Cancel"); Assert.IsFalse(string.IsNullOrEmpty(translated));
                    if (culture == "fr-FR") Assert.AreEqual("Annuler", translated);
                }
            }
            finally { UpdateText.Culture = original; }
        }
    }
}