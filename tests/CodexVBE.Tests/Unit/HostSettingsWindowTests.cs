using System;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.Serialization;
using System.Windows.Forms;
using CodexVBE;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Unit
{
    [TestClass]
    [TestCategory("Unit")]
    public sealed class HostSettingsWindowTests
    {
        private const BindingFlags Private = BindingFlags.NonPublic | BindingFlags.Instance;

        private static T Field<T>(object owner, string name)
        {
            var field = owner.GetType().GetField(name, Private);
            Assert.IsNotNull(field, name);
            return (T)field.GetValue(owner);
        }

        [TestMethod]
        [STATestMethod]
        public void SettingsKeepsSeparateEndpointAndKeyDraftsWhenProviderChanges()
        {
            var settings = new LlmSettings { ProviderName = "OpenAI API" };
            using (var window = new LlmSettingsWindow(settings))
            {
                var provider = Field<ComboBox>(window, "provider");
                Assert.AreEqual("OpenAI API", ((LlmProvider)provider.SelectedItem).Name);
                Field<TextBox>(window, "openAiEndpoint").Text = "https://example.com/v1/chat/completions";
                Field<TextBox>(window, "openAiKey").Text = "draft-only-key";
                provider.SelectedItem = Array.Find(LlmProvider.All, item => item.Name == "Ollama");
                Field<TextBox>(window, "ollamaEndpoint").Text = "http://localhost:11434/v1/chat/completions";
                provider.SelectedItem = Array.Find(LlmProvider.All, item => item.Name == "OpenAI API");
                Assert.AreEqual("https://example.com/v1/chat/completions",
                    Field<TextBox>(window, "openAiEndpoint").Text);
                Assert.AreEqual("draft-only-key", Field<TextBox>(window, "openAiKey").Text);
                Assert.AreEqual("OpenAI API", settings.ProviderName);
                Assert.AreEqual("http://localhost:11434/v1/chat/completions",
                    Field<Dictionary<string, string>>(window, "endpointDrafts")["Ollama"]);
            }
        }

        [TestMethod]
        public void EndpointValidationAcceptsHttpsAndLocalHttpOnly()
        {
            var method = typeof(LlmSettingsWindow).GetMethod("ValidateEndpoint",
                BindingFlags.NonPublic | BindingFlags.Static);
            Assert.IsNotNull(method);
            method.Invoke(null, new object[] { "" });
            method.Invoke(null, new object[] { "https://example.com/v1/chat/completions" });
            method.Invoke(null, new object[] { "http://localhost:11434/v1/chat/completions" });
            foreach (string invalid in new[] { "relative/path", "http://example.com/v1", "ftp://localhost/file" })
            {
                var error = Assert.ThrowsException<TargetInvocationException>(() =>
                    method.Invoke(null, new object[] { invalid }));
                Assert.IsInstanceOfType(error.InnerException, typeof(ArgumentException));
            }
        }

        [TestMethod]
        [STATestMethod]
        public void ProviderRowsAndUnsavedManualModelDraftsFollowTheCurrentProvider()
        {
            using (var window = new LlmSettingsWindow(new LlmSettings { ProviderName = "OpenAI API" }))
            {
                var picker = Field<ComboBox>(window, "provider");
                var grid = Field<TableLayoutPanel>(window, "grid");
                picker.SelectedItem = Array.Find(LlmProvider.All, item => item.Name == "Ollama");
                Assert.AreEqual(SizeType.Absolute, grid.RowStyles[3].SizeType);
                Assert.AreEqual(SizeType.AutoSize, grid.RowStyles[4].SizeType);
                picker.SelectedItem = Array.Find(LlmProvider.All, item => item.IsCustom);
                Assert.AreEqual(SizeType.AutoSize, grid.RowStyles[9].SizeType);
                Assert.AreEqual(SizeType.AutoSize, grid.RowStyles[10].SizeType);
                Field<TextBox>(window, "manualModels").Text = "custom-model";
                Field<CheckBox>(window, "clearKey").Checked = true;
                picker.SelectedItem = Array.Find(LlmProvider.All, item => item.IsAzure);
                Assert.AreEqual(SizeType.AutoSize, grid.RowStyles[11].SizeType);
                picker.SelectedItem = Array.Find(LlmProvider.All, item => item.IsCustom);
                Assert.AreEqual("custom-model", Field<TextBox>(window, "manualModels").Text);
                Assert.IsTrue(Field<CheckBox>(window, "clearKey").Checked);
            }
        }

        [TestMethod]
        public void MenuDiscoveryIgnoresAcceleratorsAndNonMenuBars()
        {
            var method = typeof(VbeMenu).GetMethod("FindMenu", BindingFlags.NonPublic | BindingFlags.Static);
            Assert.IsNotNull(method);
            var host = new FakeMenus();
            host.CommandBars.Add(new FakeBar { Type = 2,
                Controls = new List<FakeMenu> { new FakeMenu { Caption = "&Affichage" } } });
            var expected = new FakeMenu { Caption = " &Affichage " };
            host.CommandBars.Add(new FakeBar { Type = 1,
                Controls = new List<FakeMenu> { expected, new FakeMenu { Caption = "&Outils" } } });
            Assert.AreSame(expected, method.Invoke(null, new object[] { host, true }));
            var error = Assert.ThrowsException<TargetInvocationException>(() =>
                method.Invoke(null, new object[] { new FakeMenus(), true }));
            Assert.IsInstanceOfType(error.InnerException, typeof(InvalidOperationException));
            var incomplete = new FakeMenus();
            incomplete.CommandBars.Add(new FakeBar { Type = 1,
                Controls = new List<FakeMenu> { new FakeMenu { Caption = "View" } } });
            Assert.ThrowsException<InvalidOperationException>(() => new VbeMenu(incomplete,
                () => { }, () => { }, () => { }));
        }

        [TestMethod]
        public void AddInMetadataAndNoOpLifecycleCallbacksRemainSafeWithoutConnection()
        {
            Assert.AreEqual("CodexVBE.AddIn", ((ProgIdAttribute)Attribute.GetCustomAttribute(
                typeof(AddIn), typeof(ProgIdAttribute))).Value);
            var addin = (AddIn)FormatterServices.GetUninitializedObject(typeof(AddIn));
            object[] custom = new object[0];
            addin.OnAddInsUpdate(ref custom);
            addin.OnStartupComplete(ref custom);
            addin.OnBeginShutdown(ref custom);
        }

        public sealed class FakeMenus { public List<FakeBar> CommandBars { get; } = new List<FakeBar>(); }
        public sealed class FakeBar { public int Type { get; set; } public List<FakeMenu> Controls { get; set; } }
        public sealed class FakeMenu { public string Caption { get; set; } }
    }
}
