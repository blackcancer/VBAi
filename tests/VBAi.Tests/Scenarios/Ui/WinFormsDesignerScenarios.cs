namespace VBAi.Tests.Unit
{
    using System;
    using System.Windows.Forms;
    using VBAi;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>Vérifie que les fenêtres fixes restent instanciables hors session IDE.</summary>
    [TestClass]
    [TestCategory("Unit")]
    public sealed partial class WinFormsDesignerTests
    {
        /// <summary>Instancie les fenêtres principales et vérifie leurs contrôles Designer essentiels.</summary>
        [TestMethod]
        [STATestMethod]
        public void FixedWindowsOpenWithoutStartingAnIdeSession()
        {
            using (var chat = new ChatWindow())
            using (var git = new GitWindow())
            using (var settings = new LlmSettingsWindow())
            using (var approval = new VbeApprovalDialog("create_module\r\nProjetTest"))
            {
                Assert.IsNotNull(chat.Controls[0]);
                Assert.IsNotNull(git.Controls[0]);
                Assert.IsNotNull(settings.Controls[0]);
                Assert.AreEqual("create_module\r\nProjetTest", Find<TextBox>(approval, "details").Text);
                Assert.AreEqual(DialogResult.Yes, Find<Button>(approval, "approve").DialogResult);
                Assert.AreEqual(DialogResult.No, Find<Button>(approval, "reject").DialogResult);
                Assert.AreEqual(approval.CancelButton, Find<Button>(approval, "reject"));
                Assert.AreEqual(156, Find<Label>(settings, "providerLabel").MinimumSize.Width);
                Assert.AreEqual(30, Find<Button>(settings, "saveButton").MinimumSize.Height);
            }
        }
    }
}
