namespace VBAi.Tests.Unit
{
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using System.Windows.Forms;
    using VBAi;

    /// <summary>Vérifie les contrôles essentiels créés par les surfaces locales de discussion.</summary>
    public sealed partial class ChatWindowStateTests
    {
        /// <summary>Vérifie la création du compositeur, du transcript et des modes par défaut.</summary>
        [TestMethod]
        [STATestMethod]
        public void ConstructorAndLocalSurfacesCreateComposerTranscriptAndModes()
        {
            using (var window = Surfaces())
            {
                Assert.AreEqual(3, Get<ComboBox>(window, "modePicker").Items.Count);
                Assert.AreEqual(ChatMode.Agent, Get<ComboBox>(window, "modePicker").SelectedItem);
                Assert.IsNotNull(Get<object>(window, "prompt"));
                Assert.IsNotNull(Get<object>(window, "conversationItems"));
                Assert.IsFalse(Get<Panel>(window, "historyPanel").Visible);
                Assert.IsFalse(Get<ProgressBar>(window, "activityBar").Visible);
            }
        }
    }
}
