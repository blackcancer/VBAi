namespace VBAi.Tests.Unit
{
    using VBAi;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    [TestClass]
    [TestCategory("Unit")]
    public sealed class VbeWindowIconsTests
    {
        [TestMethod]
        public void NativeEmbeddedImagesAndMissingNamesReturnOwnedCopiesOrNull()
        {
            Assert.IsNull(VbeWindowIcons.Icon("missing-coverage-resource"));
            Assert.IsNull(VbeWindowIcons.Image("missing-coverage-resource"));
            foreach (string name in new[] { "assistant", "github", "settings" })
                using (var icon = VbeWindowIcons.Icon(name)) { Assert.IsNotNull(icon); Assert.IsTrue(icon.Width > 0); }
            foreach (string name in new[] { "github", "settings" })
                using (var image = VbeWindowIcons.Image(name)) { Assert.IsNotNull(image); Assert.IsTrue(image.Width > 0); }
        }
    }
}
