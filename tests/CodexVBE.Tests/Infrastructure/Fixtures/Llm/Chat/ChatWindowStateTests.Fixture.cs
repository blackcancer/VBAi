namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.IO;
    using System.Net;
    using System.Net.Http;
    using System.Reflection;
    using System.Threading;
    using System.Threading.Tasks;
    using System.Web.Script.Serialization;
    using System.Windows.Forms;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    public sealed partial class ChatWindowStateTests
    {
        private const BindingFlags Fields = BindingFlags.Instance | BindingFlags.NonPublic;
        private const BindingFlags Methods = BindingFlags.Instance | BindingFlags.Static | BindingFlags.NonPublic;
        private static T Get<T>(ChatWindow window, string field)
        {
            var info = typeof(ChatWindow).GetField(field, Fields);
            Assert.IsNotNull(info, "Missing ChatWindow field " + field);
            return (T)info.GetValue(window);
        }

        private static void Set(ChatWindow window, string field, object value)
        {
            var info = typeof(ChatWindow).GetField(field, Fields);
            Assert.IsNotNull(info, "Missing ChatWindow field " + field);
            info.SetValue(window, value);
        }

        private static object Call(ChatWindow window, string method, params object[] args)
        {
            var info = typeof(ChatWindow).GetMethod(method, Methods);
            Assert.IsNotNull(info, "Missing ChatWindow method " + method);
            return info.Invoke(window, args);
        }

        private static ChatWindow Surfaces()
        {
            var window = new ChatWindow();
            Call(window, "InitializeShell");
            Call(window, "InitializeComposer", new object[] { null });
            Call(window, "InitializeTranscript");
            return window;
        }

        private static ChatWindow ReadyCodexWindow(ChatSessionState session)
        {
            var window = Surfaces();
            var settings = new LlmSettings();
            Set(window, "settings", settings);
            Set(window, "tools", new LlmVbeTools(null, window, settings));
            Set(window, "currentSession", session);
            var providers = Get<ComboBox>(window, "providerPicker");
            providers.Items.Add(LlmProvider.All[0]);
            providers.SelectedIndex = 0;
            var models = Get<ComboBox>(window, "modelPicker");
            models.Items.Add(new LlmModelOption("gpt-test", "Test", true, null, new LlmEffortOption[0]));
            models.SelectedIndex = 0;
            return window;
        }

        private static ChatWindow ReadyHttpWindow(ChatSessionState session)
        {
            var window = Surfaces();
            var settings = new LlmSettings();
            Set(window, "settings", settings);
            Set(window, "tools", new LlmVbeTools(new VbeSession(new object ()), window, settings));
            Set(window, "currentSession", session);
            var providers = Get<ComboBox>(window, "providerPicker");
            providers.Items.Add(LlmProvider.All[2]);
            providers.SelectedIndex = 0;
            var models = Get<ComboBox>(window, "modelPicker");
            models.Items.Add(new LlmModelOption("local-test", "Local test"));
            models.SelectedIndex = 0;
            return window;
        }

        private sealed class ChatResponseHandler : HttpMessageHandler
        {
            private readonly Queue<string> responses = new Queue<string>();
            public readonly List<string> Requests = new List<string>();
            public ChatResponseHandler(params string[] bodies)
            {
                foreach (string body in bodies)
                    responses.Enqueue(body);
            }

            protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
            {
                Requests.Add(await request.Content.ReadAsStringAsync());
                if (responses.Count == 0)
                    throw new InvalidOperationException("Unexpected provider request.");
                return new HttpResponseMessage(HttpStatusCode.OK)
                {
                    Content = new StringContent(responses.Dequeue())
                };
            }
        }

        private static void Question(ChatWindow window, string text)
        {
            var prompt = Get<object>(window, "prompt");
            prompt.GetType().GetProperty("Text").SetValue(prompt, text, null);
        }

        private static object AddScope(ChatWindow window, string key)
        {
            var type = typeof(ChatWindow).GetNestedType("MacroScope", BindingFlags.NonPublic);
            var scope = Activator.CreateInstance(type, true);
            type.GetField("Key").SetValue(scope, key);
            type.GetField("Project").SetValue(scope, key);
            type.GetField("Name").SetValue(scope, key);
            type.GetField("Label").SetValue(scope, key);
            Get<ComboBox>(window, "scopePicker").Items.Add(scope);
            return scope;
        }

        private static void CompleteOnSta(Task task)
        {
            var deadline = DateTime.UtcNow.AddSeconds(5);
            while (!task.IsCompleted && DateTime.UtcNow < deadline)
            {
                Application.DoEvents();
                Thread.Sleep(1);
            }

            Assert.IsTrue(task.IsCompleted, "The chat operation did not complete on the STA thread.");
            task.GetAwaiter().GetResult();
        }
    }
}
