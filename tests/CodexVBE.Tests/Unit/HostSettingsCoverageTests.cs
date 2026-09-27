using System;
using System.Reflection;
using System.Windows.Forms;
using CodexVBE;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Unit
{
    [TestClass]
    [TestCategory("Unit")]
    public sealed class HostSettingsCoverageTests
    {
        private const BindingFlags PrivateInstance = BindingFlags.Instance | BindingFlags.NonPublic;

        private static T Field<T>(object instance, string name)
        {
            var field = instance.GetType().GetField(name, PrivateInstance);
            Assert.IsNotNull(field, name);
            return (T)field.GetValue(instance);
        }

        private static void Set(object instance, string name, object value)
        {
            var field = instance.GetType().GetField(name, PrivateInstance);
            Assert.IsNotNull(field, name);
            field.SetValue(instance, value);
        }

        private static object Call(object instance, string name, params object[] arguments)
        {
            var method = instance.GetType().GetMethod(name, PrivateInstance);
            Assert.IsNotNull(method, name);
            return method.Invoke(instance, arguments);
        }

        [TestMethod]
        public void MenuLookupSkipsInvalidBarsAndNormalizesLocalizedCaptions()
        {
            var host = new FakeHost { CommandBars = new object[] {
                new InvalidBar(),
                new FakeBar { Type = 2, Controls = new object[] { new FakeControl { Caption = "&Outils" } } },
                new FakeBar { Type = 1, Controls = new object[] {
                    new InvalidControl(), new FakeControl { Caption = " &Affichage " },
                    new FakeControl { Caption = "&Outils" }
                } }
            } };
            var method = typeof(VbeMenu).GetMethod("FindMenu", BindingFlags.Static | BindingFlags.NonPublic);
            Assert.IsNotNull(method);
            Assert.AreSame(((FakeBar)host.CommandBars[2]).Controls[1], method.Invoke(null, new object[] { host, new[] { "affichage", "view" } }));
            Assert.AreSame(((FakeBar)host.CommandBars[2]).Controls[2], method.Invoke(null, new object[] { host, new[] { "outils", "tools" } }));
            var missing = Assert.ThrowsException<TargetInvocationException>(() =>
                method.Invoke(null, new object[] { host, new[] { "missing" } }));
            StringAssert.Contains(missing.InnerException.Message, "VBE menu not found: missing");
        }

        [TestMethod]
        [STATestMethod]
        public void AddInShutdownClosesNativeWindowAndCanRunTwice()
        {
            var addIn = new AddIn();
            var native = new FakeNativeWindow();
            var dispatcher = new Control();
            Set(addIn, "nativeChatWindow", native);
            Set(addIn, "dispatcher", dispatcher);
            object[] custom = null;
            addIn.OnBeginShutdown(ref custom);
            Assert.AreEqual(1, native.CloseCount);
            Assert.IsTrue(dispatcher.IsDisposed);
            Assert.IsNull(Field<object>(addIn, "nativeChatWindow"));
            Assert.IsNull(Field<object>(addIn, "dispatcher"));
            addIn.OnDisconnection(0, ref custom);
            Assert.AreEqual(1, native.CloseCount);
        }

        [TestMethod]
        [STATestMethod]
        public void SettingsKeepSeparateDraftsWhenSwitchingProviders()
        {
            var settings = new LlmSettings { ProviderName = "OpenAI API", OpenAiEndpoint = "https://original.example/v1" };
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
            foreach (var valid in new[] { "", "  ", "https://api.example/v1", "http://localhost:11434", "http://127.0.0.1:1234" })
                method.Invoke(null, new object[] { valid });
            foreach (var invalid in new[] { "not a url", "http://api.example/v1", "ftp://localhost/data" })
            {
                var thrown = Assert.ThrowsException<TargetInvocationException>(() => method.Invoke(null, new object[] { invalid }));
                Assert.IsInstanceOfType(thrown.InnerException, typeof(ArgumentException));
            }
        }

        public sealed class FakeHost { public object[] CommandBars { get; set; } }
        public sealed class FakeBar { public int Type { get; set; } public object[] Controls { get; set; } }
        public sealed class InvalidBar { public int Type { get { throw new InvalidOperationException("no type"); } } }
        public sealed class FakeControl { public string Caption { get; set; } }
        public sealed class InvalidControl { public string Caption { get { throw new InvalidOperationException("no caption"); } } }
        public sealed class FakeNativeWindow { public int CloseCount { get; private set; } public void Close() { CloseCount++; } }
    }
}
