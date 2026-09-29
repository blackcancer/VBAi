namespace VBAi.Tests.Unit
{
    using System;
    using System.Reflection;
    using System.Windows.Forms;
    using VBAi;
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

namespace VBAi.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using System.Reflection;
    using System.Runtime.InteropServices;
    using System.Runtime.Serialization;
    using System.Windows.Forms;
    using VBAi;
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
                Assert.AreEqual(SizeType.AutoSize, grid.RowStyles[3].SizeType);
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

namespace VBAi.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Threading.Tasks;
    using System.Windows.Forms;
    using VBAi;
    using VBAi.Tests.Infrastructure;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    [TestClass, TestCategory("Unit")]
    public sealed class LlmSettingsWindowBoundaryTests
    {
        private static T Get<T>(LlmSettingsWindow window, string field) { return LlmBoundaryScope.Get<T>(window, field); }
        private static void Call(LlmSettingsWindow window, string method, params object[] args) { LlmBoundaryScope.Call(window, method, args); }
        private static Task Refresh(LlmSettingsWindow window, string method, params object[] args) { return (Task)LlmBoundaryScope.Call(window, method, args); }
        private static void Select(LlmSettingsWindow window, string name) { Get<ComboBox>(window, "provider").SelectedItem = LlmBoundaryScope.Provider(name); }
        [STATestMethod]
        public void DesignerThemeAndLayoutEventsRestoreGlobalThemeAndCoverPartialCleanup()
        {
            using (var scope = new LlmBoundaryScope())
            {
                foreach (bool configured in new[] { false, true }) using (var window = configured ? scope.Window(new LlmSettings { ProviderName = "OpenAI API" }) : new LlmSettingsWindow())
                {
                    var themes = Get<ComboBox>(window, "themePicker"); themes.SelectedIndex = -1; themes.SelectedIndex = 1; LlmSettingsWindow.SelectTheme = c => { throw new IOException("theme fixture failure"); }; themes.SelectedIndex = 2; Assert.IsTrue(scope.Notices.Contains("theme fixture failure")); LlmSettingsWindow.SelectTheme = UiTheme.Select;
                    LayoutEventHandler reenter = (s, e) => Call(window, "FitContentHeight"); window.Layout += reenter; try { window.Show(); Call(window, "FitContentHeight"); Call(window, "OnShown", EventArgs.Empty); Assert.IsTrue(window.ClientSize.Height > 0); } finally { window.Layout -= reenter; }
                    if (!configured) { Call(window, "UpdateRows"); LlmBoundaryScope.Pump(Refresh(window, "RefreshGitHubAsync", false)); }
                    Call(window, "Dispose", false); window.Dispose(); window.Dispose(); Call(window, "FitContentHeight");
                }
                using (var window = scope.Window(new LlmSettings { ProviderName = "OpenAI API" })) { var tooltip = Get<ToolTip>(window, "githubToolTips"); try { LlmBoundaryScope.Set(window, "githubToolTips", null); Call(window, "Dispose", true); } finally { LlmBoundaryScope.Set(window, "githubToolTips", tooltip); tooltip.Dispose(); } }
            }
        }
        [STATestMethod]
        public void EveryProviderRowDraftLoginAndThemeCallbackUsesOnlyIsolatedBoundaries()
        {
            using (var scope = new LlmBoundaryScope())
            using (var window = scope.Window(new LlmSettings { ProviderName = "Unknown", CustomProviderName = null, GitHubAccount = "saved-user", VbeEditApproval = "AskEachTime" }))
            {
                Assert.AreEqual("Codex", ((LlmProvider)Get<ComboBox>(window, "provider").SelectedItem).Name); Assert.AreEqual(1, Get<ComboBox>(window, "approvalPicker").SelectedIndex); window.Show();
                foreach (var provider in LlmProvider.All)
                {
                    Select(window, provider.Name); Call(window, "UpdateRows"); Get<TextBox>(window, provider.Local ? "ollamaEndpoint" : "openAiEndpoint").Text = " https://fixture.invalid/v1/chat/completions "; Get<TextBox>(window, "openAiKey").Text = " fixture-key "; Get<CheckBox>(window, "clearKey").Checked = true; Get<TextBox>(window, "manualModels").Text = " model-a "; Call(window, "CaptureDraft"); Get<CheckBox>(window, "clearKey").Checked = false; Call(window, "CaptureDraft");
                    if (provider.IsCodex || provider.IsCopilot) { LlmBoundaryScope.Click(Get<Button>(window, "codexLogin")); LlmBoundaryScope.Click(Get<Button>(window, "codexRefresh")); Assert.IsTrue(Get<Label>(window, "codexStatus").Text.Contains("fixture")); }
                }
                Assert.IsTrue(scope.Logins.Contains("copilot")); Assert.IsTrue(scope.Logins.Contains("codex")); Get<ComboBox>(window, "provider").SelectedIndex = -1; Call(window, "UpdateRows"); Select(window, "Codex"); LlmSettingsWindow.StartCodexLogin = () => { throw new IOException("login fixture failure"); }; LlmBoundaryScope.Click(Get<Button>(window, "codexLogin")); Assert.AreEqual("login fixture failure", Get<Label>(window, "codexStatus").Text);
            }
        }
        [STATestMethod]
        public void SavingValidatedDraftsKeysApprovalAndManualModelsNeverWritesUserSettings()
        {
            using (var scope = new LlmBoundaryScope())
            {
                foreach (var policy in new[] { "Automatic", "AskEachTime", "ReadOnly" })
                {
                    var settings = new LlmSettings { ProviderName = "OpenAI API", VbeEditApproval = policy, ManualModelLists = null }; settings.SetKey(LlmBoundaryScope.Provider("OpenAI API"), "old-fixture-key"); using (var window = scope.Window(settings))
                    {
                        Get<ComboBox>(window, "approvalPicker").SelectedIndex = policy == "ReadOnly" ? 2 : policy == "AskEachTime" ? 1 : 0; Get<TextBox>(window, "openAiEndpoint").Text = "https://fixture.invalid/v1"; Get<CheckBox>(window, "clearKey").Checked = true; Get<TextBox>(window, "openAiKey").Text = policy == "Automatic" ? " new-fixture-key " : " "; Select(window, "Personnalisé (OpenAI)"); Get<TextBox>(window, "openAiEndpoint").Text = "https://custom.invalid/v1/chat/completions"; Get<TextBox>(window, "manualModels").Text = " first\nsecond "; Get<TextBox>(window, "customName").Text = " Friendly "; Get<CheckBox>(window, "azureEntra").Checked = true;
                        if (policy == "ReadOnly") { Get<ComboBox>(window, "githubAccount").Items.Add("fixture-account"); Get<ComboBox>(window, "githubAccount").SelectedItem = "fixture-account"; }
                        LlmBoundaryScope.Click(Get<Button>(window, "saveButton")); Assert.AreEqual(DialogResult.OK, window.DialogResult); Assert.AreEqual(policy, settings.VbeEditApproval); Assert.AreEqual("Friendly", settings.CustomProviderName); Assert.AreEqual("first\nsecond", settings.ManualModelLists["Personnalisé (OpenAI)"]); Assert.AreEqual(policy == "ReadOnly" ? "fixture-account" : null, settings.GitHubAccount); Assert.AreEqual(policy == "Automatic" ? "new-fixture-key" : null, settings.GetKey(LlmBoundaryScope.Provider("OpenAI API")));
                    }
                }
                Assert.AreEqual(3, scope.Saves);
                using (var window = scope.Window(new LlmSettings { ProviderName = "OpenAI API" })) { Get<TextBox>(window, "openAiEndpoint").Text = "http://remote.invalid"; Call(window, "Save"); Assert.AreEqual(DialogResult.None, window.DialogResult); Assert.AreEqual(3, scope.Saves); Assert.IsTrue(scope.Notices.Last().Contains("HTTPS")); Get<TextBox>(window, "openAiEndpoint").Text = "https://valid.invalid"; LlmSettingsWindow.WriteSettings = s => { throw new IOException("save fixture failure"); }; Call(window, "Save"); Assert.AreEqual("save fixture failure", scope.Notices.Last()); }
            }
        }
        [STATestMethod]
        public void NativeVbeThemeIsExplicitPersistedAndRolledBackWhenApplicationFails()
        {
            using (var scope = new LlmBoundaryScope())
            {
                var applied = new System.Collections.Generic.List<bool>();
                LlmSettingsWindow.SelectNativeVbeTheme = value => applied.Add(value);
                var settings = new LlmSettings { ProviderName = "OpenAI API", NativeVbeDarkTheme = true };
                using (var window = scope.Window(settings))
                {
                    var toggle = Get<CheckBox>(window, "nativeVbeDark");
                    Assert.IsTrue(toggle.Checked);
                    toggle.Checked = false;
                    Call(window, "Save");
                    Assert.IsFalse(settings.NativeVbeDarkTheme);
                    CollectionAssert.AreEqual(new[] { false }, applied);
                }
                LlmSettingsWindow.SelectNativeVbeTheme = value => { if (value) throw new IOException("native theme unavailable"); };
                using (var window = scope.Window(settings))
                {
                    Get<CheckBox>(window, "nativeVbeDark").Checked = true;
                    Call(window, "Save");
                    Assert.IsFalse(settings.NativeVbeDarkTheme);
                    Assert.AreEqual("native theme unavailable", scope.Notices.Last());
                }
            }
        }
        [STATestMethod]
        public void GithubLoginAccountSelectionAndConcurrentDiscoveryRespectUiLifetime()
        {
            using (var scope = new LlmBoundaryScope())
            {
                foreach (bool login in new[] { false, true }) foreach (var accounts in new[] { "", "one", "one\ntwo" }) foreach (var selected in new[] { null, "missing" })
                    using (var window = scope.Window(new LlmSettings { ProviderName = "OpenAI API", GitHubAccount = selected }))
                    {
                        int commands = 0; LlmBoundaryScope.Set(window, "githubService", new GitHubAccountService((command, token) => { commands++; return Task.FromResult(command.Contains(" login ") ? "" : accounts); })); window.Show(); LlmBoundaryScope.Pump(Refresh(window, "RefreshGitHubAsync", login)); Assert.IsTrue(Get<Button>(window, "saveButton").Enabled); var picker = Get<ComboBox>(window, "githubAccount"); Assert.AreEqual(selected ?? (login && accounts == "one" ? "one" : UiText.Get("Automatic Git selection")), picker.SelectedItem); Assert.IsTrue(commands > 0); LlmBoundaryScope.Click(Get<Button>(window, "githubRefresh")); LlmBoundaryScope.Click(Get<Button>(window, "githubLogin"));
                    }
                using (var window = scope.Window(new LlmSettings { ProviderName = "OpenAI API" })) { var pending = new TaskCompletionSource<string>(); LlmBoundaryScope.Set(window, "githubService", new GitHubAccountService((c, t) => pending.Task)); var first = Refresh(window, "RefreshGitHubAsync", false); LlmBoundaryScope.Pump(Refresh(window, "RefreshGitHubAsync", true)); Assert.IsTrue(Get<bool>(window, "githubBusy")); Assert.IsFalse(Get<Button>(window, "saveButton").Enabled); pending.SetResult("fixture-user"); LlmBoundaryScope.Pump(first); Assert.IsFalse(Get<bool>(window, "githubBusy")); }
                foreach (bool disposed in new[] { false, true }) foreach (bool canceled in new[] { false, true })
                    using (var window = scope.Window(new LlmSettings { ProviderName = "OpenAI API" })) { var pending = new TaskCompletionSource<string>(); LlmBoundaryScope.Set(window, "githubService", new GitHubAccountService((c, t) => pending.Task)); var refresh = Refresh(window, "RefreshGitHubAsync", false); if (disposed) window.Dispose(); if (canceled) pending.SetCanceled(); else pending.SetException(new IOException("github fixture failure")); LlmBoundaryScope.Pump(refresh); if (!disposed && !canceled) Assert.AreEqual("github fixture failure", Get<Label>(window, "githubStatus").Text); }
                using (var window = scope.Window(new LlmSettings { ProviderName = "OpenAI API" })) { var pending = new TaskCompletionSource<string>(); LlmBoundaryScope.Set(window, "githubService", new GitHubAccountService((c, t) => pending.Task)); var refresh = Refresh(window, "RefreshGitHubAsync", false); window.Dispose(); pending.SetResult("late"); LlmBoundaryScope.Pump(refresh); Assert.IsFalse(Get<bool>(window, "githubBusy")); }
            }
        }
        [STATestMethod]
        public void CodexAndCopilotStatusResultsErrorsAndLateRepliesRespectProviderChanges()
        {
            using (var scope = new LlmBoundaryScope())
            {
                foreach (var provider in new[] { "Codex", "GitHub Copilot" }) foreach (bool failure in new[] { false, true }) foreach (var lifetime in new[] { "current", "changed", "disposed" })
                    using (var window = scope.Window(new LlmSettings { ProviderName = provider }))
                    {
                        window.Show(); var copilot = new TaskCompletionSource<string>(); var codex = new TaskCompletionSource<CodexAccountStatus>(); LlmSettingsWindow.ReadCopilotStatus = () => copilot.Task; LlmSettingsWindow.ReadCodexStatus = () => codex.Task; var refresh = Refresh(window, "RefreshCodexStatusAsync"); if (lifetime == "changed") Select(window, "Ollama"); if (lifetime == "disposed") window.Dispose(); if (failure) { if (provider == "Codex") codex.SetException(new IOException("late fixture failure")); else copilot.SetException(new IOException("late fixture failure")); } else { if (provider == "Codex") codex.SetResult(new CodexAccountStatus(false, "late fixture connected")); else copilot.SetResult("late fixture connected"); }
                        LlmBoundaryScope.Pump(refresh); if (lifetime == "current") { Assert.AreEqual(failure ? "late fixture failure" : "late fixture connected", Get<Label>(window, "codexStatus").Text); if (provider == "Codex") Assert.IsTrue(Get<Button>(window, "codexLogin").Enabled); } else Assert.AreNotEqual(failure ? "late fixture failure" : "late fixture connected", Get<Label>(window, "codexStatus").Text);
                        LlmSettingsWindow.ReadCopilotStatus = () => Task.FromResult("Copilot fixture connected"); LlmSettingsWindow.ReadCodexStatus = () => Task.FromResult(new CodexAccountStatus(true, "ChatGPT fixture connected"));
                    }
            }
        }
    }
}
