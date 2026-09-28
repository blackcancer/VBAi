using System;
using System.ComponentModel;
using System.Linq;
using System.Windows.Forms;
using CodexVBE;
using CodexVBE.Tests.Infrastructure;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Unit
{
    /// <summary>Vérifie les métadonnées, les actions et l’intégration d’À propos.</summary>
    [TestClass, TestCategory("Unit")]
    public sealed class AboutWindowTests
    {
        [STATestMethod]
        public void AllSupportedLanguagesKeepVersionLinksAndDesignerControlsUsable()
        {
            using (var theme = new ThemeScope())
            using (var culture = new LocalizationScope())
                foreach (var language in UiLanguages.All)
                {
                    LocalizationScope.Set(language.CultureName);
                    using (var window = new AboutWindow())
                    {
                        Assert.AreEqual(UiText.Get("About VBAi"), window.Text);
                        Assert.IsNotNull(UiInvoke.Field<PictureBox>(window, "brandImage").Image);
                        Assert.AreEqual("VBAi", UiInvoke.Field<Label>(window, "productName").Text);
                        Assert.AreEqual("Your AI agent for VBA", UiInvoke.Field<Label>(window, "tagline").Text);
                        Assert.IsTrue(window.TechnicalDetails.Contains(UpdateState.ProductVersion));
                        Assert.IsTrue(window.TechnicalDetails.Contains(".NET Framework 4.8"));
                        Assert.IsFalse(window.TechnicalDetails.Contains(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile)));
                        Assert.AreEqual(UiText.Get("Close"), ((Button)window.CancelButton).Text);
                        Assert.AreSame(window.AcceptButton, window.CancelButton);
                        foreach (var name in new[] { "projectLink", "documentationLink", "issuesLink" })
                            Assert.IsTrue(SafeLinks.Allowed((string)UiInvoke.Field<LinkLabel>(window, name).Tag));
                        Assert.AreEqual(language.CultureName == "ar-SA", window.RightToLeftLayout);
                        Assert.AreEqual(RightToLeft.No, UiInvoke.Field<Label>(window, "platformValue").RightToLeft);
                    }
                }
        }

        [STATestMethod]
        public void CopyAndResourceActionsReportFailuresWithoutOpeningAnythingDuringConstruction()
        {
            using (var culture = new LocalizationScope())
            using (var window = new AboutWindow())
            {
                string copied = null, opened = null;
                window.CopyText = text => copied = text;
                window.OpenLink = url => opened = url;
                UiInvoke.Call(typeof(AboutWindow), "CopyDetails_Click", window, null, EventArgs.Empty);
                Assert.AreEqual(window.TechnicalDetails, copied);
                Assert.AreEqual(UiText.Get("Technical details copied."), UiInvoke.Field<Label>(window, "status").Text);
                var project = UiInvoke.Field<LinkLabel>(window, "projectLink");
                UiInvoke.Call(typeof(AboutWindow), "ResourceLink_Click", window, project, null);
                Assert.AreEqual("https://github.com/blackcancer/CodexVBE", opened);
                Assert.AreEqual("", UiInvoke.Field<Label>(window, "status").Text);
                window.CopyText = text => throw new InvalidOperationException("Clipboard unavailable");
                UiInvoke.Call(typeof(AboutWindow), "CopyDetails_Click", window, null, EventArgs.Empty);
                Assert.AreEqual(UiText.Get("Unable to copy technical details."), UiInvoke.Field<Label>(window, "status").Text);
                window.OpenLink = url => throw new InvalidOperationException("Browser unavailable");
                UiInvoke.Call(typeof(AboutWindow), "ResourceLink_Click", window, project, null);
                Assert.AreEqual(UiText.Get("Unable to open the link."), UiInvoke.Field<Label>(window, "status").Text);
            }
        }

        [STATestMethod]
        public void DesignerConstructorAndHostDescriptionDoNotRequireAVbeSession()
        {
            using (var window = (AboutWindow)LicenseManager.CreateWithContext(typeof(AboutWindow), new DesignContext()))
            {
                Assert.AreEqual("About VBAi", window.Text);
                Assert.AreEqual("Visual Basic Editor", UiInvoke.Field<Label>(window, "hostValue").Text);
                Assert.IsNotNull(UiInvoke.Field<PictureBox>(window, "brandImage").Image);
            }
            Assert.AreEqual("Microsoft Excel · Visual Basic Editor", AboutWindow.HostDescription("excel"));
            Assert.AreEqual("SOLIDWORKS · Visual Basic Editor", AboutWindow.HostDescription("SLDWORKS"));
            Assert.AreEqual("other-host", AboutWindow.HostDescription("other-host"));
        }

        [STATestMethod]
        public void AboutOpensTheUpdateConfigurationWithItsOwner()
        {
            var original = AddIn.ShowModal;
            try
            {
                using (var window = new AboutWindow())
                {
                    int shown = 0;
                    AddIn.ShowModal = (dialog, owner) => { Assert.IsInstanceOfType(dialog, typeof(UpdateWindow)); Assert.AreSame(window, owner); shown++; return DialogResult.Cancel; };
                    window.Show();
                    UiInvoke.Field<Button>(window, "updates").PerformClick(); Assert.AreEqual(1, shown);
                }
            }
            finally { AddIn.ShowModal = original; }
        }

        [STATestMethod]
        public void AboutOpensFromChatAndVbeWithoutProviderConfiguration()
        {
            var chatShow = ChatWindow.ShowModal;
            var hostShow = AddIn.ShowModal;
            int shown = 0;
            try
            {
                using (var chat = new ChatWindow())
                {
                    ChatWindow.ShowModal = (form, owner) => { Assert.IsInstanceOfType(form, typeof(AboutWindow)); Assert.AreSame(chat, owner); shown++; return DialogResult.OK; };
                    UiInvoke.Field<ToolStripMenuItem>(chat, "about").PerformClick();
                }
                AddIn.ShowModal = (form, owner) => { Assert.IsInstanceOfType(form, typeof(AboutWindow)); Assert.AreEqual(new IntPtr(123), owner.Handle); shown++; return DialogResult.OK; };
                AboutWindow.ShowForVbe(new AboutHost());
                Assert.AreEqual(2, shown);
            }
            finally { ChatWindow.ShowModal = chatShow; AddIn.ShowModal = hostShow; }
        }
        [STATestMethod]
        public void OwnedMetadataAppearanceAndOwnerFailuresRespectEveryRuntimeBoundary()
        {
            using (var metadata = new TechnicalUiMetadataScope())
            using (var theme = new ThemeScope())
            {
                AboutWindow.MetadataAssembly = TechnicalUiMetadataScope.WithoutInformation; AboutWindow.ProcessIs64Bit = () => false; AboutWindow.RuntimeVersion = () => null;
                using (var window = new AboutWindow())
                {
                    StringAssert.Contains(window.TechnicalDetails, "5.6.7.8"); StringAssert.Contains(window.TechnicalDetails, "x86");
                    Assert.IsNotNull(AboutWindow.ResolveImageReader(null, new System.ResolveEventArgs("System.Resources.Extensions, Version=4.0.0.0, Culture=neutral, PublicKeyToken=cc7b13ffcd2ddd51")));
                    Assert.IsNull(AboutWindow.ResolveImageReader(null, new System.ResolveEventArgs("Unrelated.Assembly")));
                    var handle = window.Handle;
                    foreach (bool contrast in new[] { false, true })
                    foreach (var choice in new[] { ThemeChoice.Light, ThemeChoice.Dark })
                    {
                        UiTheme.HighContrast = () => contrast; ThemeScope.SetChoice(choice);
                        ModernEditorDebugFixture.Wait(System.Threading.Tasks.Task.Run(() => UiInvoke.Call(typeof(AboutWindow), "ApplyAppearance", window)));
                        Assert.AreEqual(contrast ? System.Drawing.SystemColors.HotTrack : choice == ThemeChoice.Dark ? System.Drawing.Color.FromArgb(147, 197, 253) : System.Drawing.Color.FromArgb(29, 78, 216), UiInvoke.Field<LinkLabel>(window, "projectLink").LinkColor);
                    }
                    window.Dispose(); UiInvoke.Call(typeof(AboutWindow), "ApplyAppearance", window);
                }
                var original = AddIn.ShowModal;
                try { int shown = 0; AddIn.ShowModal = (form, owner) => { Assert.IsNull(owner); shown++; return DialogResult.Cancel; }; AboutWindow.ShowForVbe(null); Assert.AreEqual(1, shown); }
                finally { AddIn.ShowModal = original; }
            }
        }

        [STATestMethod]
        public void DesignerDisposalHandlesAbsentImageAndAlreadyReleasedComponents()
        {
            using (var window = (AboutWindow)LicenseManager.CreateWithContext(typeof(AboutWindow), new DesignContext()))
            {
                var image = UiInvoke.Field<PictureBox>(window, "brandImage"); image.Image.Dispose(); image.Image = null;
                var components = UiInvoke.Field<System.ComponentModel.IContainer>(window, "components");
                typeof(AboutWindow).GetField("components", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic).SetValue(window, null);
                try { UiInvoke.Call(typeof(AboutWindow), "Dispose", window, false); window.Dispose(); }
                finally { components.Dispose(); }
            }
        }
        public sealed class AboutHost { public AboutMainWindow MainWindow { get; } = new AboutMainWindow(); }
        public sealed class AboutMainWindow { public long HWnd => 123; }
    }
}
