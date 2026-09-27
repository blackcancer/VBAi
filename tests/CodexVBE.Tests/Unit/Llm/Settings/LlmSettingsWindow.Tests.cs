namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Reflection;
    using System.Windows.Forms;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    public sealed partial class HostSettingsCoverageTests
    {
        [TestMethod]
        [STATestMethod]
        public void SettingsKeepSeparateDraftsWhenSwitchingProviders()
        {
            var settings = new LlmSettings
            {
                ProviderName = "OpenAI API",
                OpenAiEndpoint = "https://original.example/v1"
            };
            using (var window = new LlmSettingsWindow(settings))
            {
                var provider = Field<ComboBox>(window, "provider");
                var endpoint = Field<TextBox>(window, "openAiEndpoint");
                var localEndpoint = Field<TextBox>(window, "ollamaEndpoint");
                var key = Field<TextBox>(window, "openAiKey");
                var clear = Field<CheckBox>(window, "clearKey");
                Assert.AreEqual("https://original.example/v1", endpoint.Text);
                endpoint.Text = " https://changed.example/v1 ";
                key.Text = " secret ";
                clear.Checked = true;
                provider.SelectedItem = Array.Find(LlmProvider.All, p => p.IsOllama);
                localEndpoint.Text = " http://localhost:11434/v1 ";
                provider.SelectedItem = Array.Find(LlmProvider.All, p => p.Name == "OpenAI API");
                Assert.AreEqual("https://changed.example/v1", endpoint.Text);
                Assert.AreEqual("secret", key.Text);
                Assert.IsTrue(clear.Checked);
                provider.SelectedItem = Array.Find(LlmProvider.All, p => p.IsOllama);
                Assert.AreEqual("http://localhost:11434/v1", localEndpoint.Text);
                Assert.AreEqual("", key.Text);
                Assert.IsFalse(clear.Checked);
            }
        }

        [TestMethod]
        public void SettingsEndpointValidationAcceptsHttpsAndLoopbackHttpOnly()
        {
            var method = typeof(LlmSettingsWindow).GetMethod("ValidateEndpoint", BindingFlags.Static | BindingFlags.NonPublic);
            Assert.IsNotNull(method);
            foreach (var valid in new[]
            {
                "",
                "  ",
                "https://api.example/v1",
                "http://localhost:11434",
                "http://127.0.0.1:1234"
            }

            )
                method.Invoke(null, new object[] { valid });
            foreach (var invalid in new[]
            {
                "not a url",
                "http://api.example/v1",
                "ftp://localhost/data"
            }

            )
            {
                var thrown = Assert.ThrowsException<TargetInvocationException>(() => method.Invoke(null, new object[] { invalid }));
                Assert.IsInstanceOfType(thrown.InnerException, typeof(ArgumentException));
            }
        }
    }
}

namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using System.Reflection;
    using System.Runtime.InteropServices;
    using System.Runtime.Serialization;
    using System.Windows.Forms;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    public sealed partial class HostSettingsWindowTests
    {
        [TestMethod]
        [STATestMethod]
        public void SettingsKeepsSeparateEndpointAndKeyDraftsWhenProviderChanges()
        {
            var settings = new LlmSettings
            {
                ProviderName = "OpenAI API"
            };
            using (var window = new LlmSettingsWindow(settings))
            {
                var provider = Field<ComboBox>(window, "provider");
                Assert.AreEqual("OpenAI API", ((LlmProvider)provider.SelectedItem).Name);
                Field<TextBox>(window, "openAiEndpoint").Text = "https://example.com/v1/chat/completions";
                Field<TextBox>(window, "openAiKey").Text = "draft-only-key";
                provider.SelectedItem = Array.Find(LlmProvider.All, item => item.Name == "Ollama");
                Field<TextBox>(window, "ollamaEndpoint").Text = "http://localhost:11434/v1/chat/completions";
                provider.SelectedItem = Array.Find(LlmProvider.All, item => item.Name == "OpenAI API");
                Assert.AreEqual("https://example.com/v1/chat/completions", Field<TextBox>(window, "openAiEndpoint").Text);
                Assert.AreEqual("draft-only-key", Field<TextBox>(window, "openAiKey").Text);
                Assert.AreEqual("OpenAI API", settings.ProviderName);
                Assert.AreEqual("http://localhost:11434/v1/chat/completions", Field<Dictionary<string, string>>(window, "endpointDrafts")["Ollama"]);
            }
        }

        [TestMethod]
        public void EndpointValidationAcceptsHttpsAndLocalHttpOnly()
        {
            var method = typeof(LlmSettingsWindow).GetMethod("ValidateEndpoint", BindingFlags.NonPublic | BindingFlags.Static);
            Assert.IsNotNull(method);
            method.Invoke(null, new object[] { "" });
            method.Invoke(null, new object[] { "https://example.com/v1/chat/completions" });
            method.Invoke(null, new object[] { "http://localhost:11434/v1/chat/completions" });
            foreach (string invalid in new[]
            {
                "relative/path",
                "http://example.com/v1",
                "ftp://localhost/file"
            }

            )
            {
                var error = Assert.ThrowsException<TargetInvocationException>(() => method.Invoke(null, new object[] { invalid }));
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
    }
}
