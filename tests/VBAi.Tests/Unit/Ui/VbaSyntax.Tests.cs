using System.Drawing;
using System.Linq;
using VBAi;
using VBAi.Tests.Infrastructure;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace VBAi.Tests.Unit
{
    /// <summary>Vérifie la tokenisation et la coloration syntaxique du code VBA.</summary>
    [TestClass]
    public sealed class VbaSyntaxTests
    {
        /// <summary>Préserve tous les caractères et distingue chaînes, commentaires et mots-clés.</summary>
        [TestMethod]
        public void TokenizationPreservesEveryCharacterAndDistinguishesStringsCommentsKeywords()
        {
            var source = "x = \"Sub \"\"quoted\"\"\" 'Then\nRem comment\nIf foo Then trailing";
            var parts=VbaSyntax.Parts(source).ToArray();
            Assert.AreEqual(source,string.Concat(parts.Select(p => p.Text)));
            CollectionAssert.AreEqual(new[] { "plain","string","plain","comment","plain","comment","plain","keyword","plain","keyword","plain" },parts.Select(p=>p.Kind).ToArray());
            Assert.AreEqual("string",VbaSyntax.Parts("\"If\"").Single().Kind);
            Assert.AreEqual("keyword",VbaSyntax.Parts("Sub").Single().Kind);
            Assert.AreEqual("plain",VbaSyntax.Parts("foo").Single().Kind);
            Assert.AreEqual(0,VbaSyntax.Parts("").Count());
        }
        /// <summary>Suit les palettes des thèmes et utilise la couleur de texte par défaut pour les catégories inconnues.</summary>
        [TestMethod]
        public void SyntaxColorsFollowBothThemePalettesAndUnknownKindsUseForeground()
        {
            using(var scope=new ThemeScope())
            {
                foreach(var dark in new[]{false,true})
                {
                    ThemeScope.SetChoice(dark ? ThemeChoice.Dark : ThemeChoice.Light);
                    Assert.AreEqual(dark ? Color.LightGreen : Color.ForestGreen,VbaSyntax.Color("comment"));
                    Assert.AreEqual(dark ? Color.SandyBrown : Color.Brown,VbaSyntax.Color("string"));
                    Assert.AreEqual(dark ? Color.LightSkyBlue : Color.RoyalBlue,VbaSyntax.Color("keyword"));
                    Assert.AreEqual(UiTheme.Foreground,VbaSyntax.Color("plain"));
                    Assert.AreEqual(UiTheme.Foreground,VbaSyntax.Color(null));
                }
            }
        }
    }
}
