namespace VBAi.Tests.Unit
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
    using VBAi;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>Fournit les constructeurs de fenêtres et aides réflexives partagés par les tests de discussion.</summary>
    public sealed partial class ChatWindowStateTests
    {
        /// <summary>Options de réflexion pour les champs d’instance privés de ChatWindow.</summary>
        private const BindingFlags Fields = BindingFlags.Instance | BindingFlags.NonPublic;
        /// <summary>Options de réflexion pour les méthodes privées statiques ou d’instance de ChatWindow.</summary>
        private const BindingFlags Methods = BindingFlags.Instance | BindingFlags.Static | BindingFlags.NonPublic;
        /// <summary>Lit un champ privé de la fenêtre de discussion.</summary>
        /// <typeparam name="T">Type attendu du champ.</typeparam>
        /// <param name="window">Fenêtre à inspecter.</param>
        /// <param name="field">Nom du champ.</param>
        /// <returns>Valeur du champ convertie en <typeparamref name="T"/>.</returns>
        private static T Get<T>(ChatWindow window, string field)
        {
            var info = typeof(ChatWindow).GetField(field, Fields);
            Assert.IsNotNull(info, "Missing ChatWindow field " + field);
            return (T)info.GetValue(window);
        }

        /// <summary>Modifie un champ privé de la fenêtre de discussion.</summary>
        /// <param name="window">Fenêtre à modifier.</param>
        /// <param name="field">Nom du champ.</param>
        /// <param name="value">Nouvelle valeur du champ.</param>
        private static void Set(ChatWindow window, string field, object value)
        {
            var info = typeof(ChatWindow).GetField(field, Fields);
            Assert.IsNotNull(info, "Missing ChatWindow field " + field);
            info.SetValue(window, value);
        }

        /// <summary>Appelle une méthode privée de la fenêtre de discussion.</summary>
        /// <param name="window">Fenêtre cible.</param>
        /// <param name="method">Nom de la méthode.</param>
        /// <param name="args">Arguments à transmettre.</param>
        /// <returns>Valeur renvoyée par la méthode.</returns>
        private static object Call(ChatWindow window, string method, params object[] args)
        {
            var info = typeof(ChatWindow).GetMethod(method, Methods);
            Assert.IsNotNull(info, "Missing ChatWindow method " + method);
            return info.Invoke(window, args);
        }

        /// <summary>Crée les trois surfaces UI sans initialiser de session VBE.</summary>
        /// <returns>Fenêtre initialisée avec shell, compositeur et transcript.</returns>
        private static ChatWindow Surfaces()
        {
            var window = new ChatWindow();
            Call(window, "InitializeShell");
            Call(window, "InitializeComposer", new object[] { null });
            Call(window, "InitializeTranscript");
            return window;
        }

        /// <summary>Crée une fenêtre prête à utiliser un fournisseur Codex simulé.</summary>
        /// <param name="session">État de session à associer à la fenêtre.</param>
        /// <returns>Fenêtre configurée avec un fournisseur et un modèle Codex de test.</returns>
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

        /// <summary>Crée une fenêtre prête à utiliser un fournisseur HTTP local simulé.</summary>
        /// <param name="session">État de session à associer à la fenêtre.</param>
        /// <returns>Fenêtre configurée avec un fournisseur et un modèle HTTP de test.</returns>
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

        /// <summary>Répond aux requêtes HTTP avec une séquence de corps JSON prédéfinie.</summary>
        private sealed class ChatResponseHandler : HttpMessageHandler
        {
            /// <summary>Réponses à renvoyer dans l’ordre des requêtes.</summary>
            private readonly Queue<string> responses = new Queue<string>();
            /// <summary>Corps des requêtes reçues.</summary>
            public readonly List<string> Requests = new List<string>();
            /// <summary>Initialise la file de réponses du faux fournisseur.</summary>
            /// <param name="bodies">Corps HTTP à renvoyer successivement.</param>
            public ChatResponseHandler(params string[] bodies)
            {
                foreach (string body in bodies)
                    responses.Enqueue(body);
            }

            /// <summary>Enregistre le corps de la requête et renvoie la prochaine réponse préparée.</summary>
            /// <param name="request">Requête HTTP du fournisseur.</param>
            /// <param name="cancellationToken">Jeton d’annulation de la requête.</param>
            /// <returns>Réponse HTTP contenant le prochain corps JSON.</returns>
            /// <exception cref="InvalidOperationException">Aucune réponse prédéfinie ne reste disponible.</exception>
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

        /// <summary>Place le texte fourni dans l’éditeur de prompt par réflexion.</summary>
        /// <param name="window">Fenêtre dont l’éditeur doit être modifié.</param>
        /// <param name="text">Texte de question à saisir.</param>
        private static void Question(ChatWindow window, string text)
        {
            var prompt = Get<object>(window, "prompt");
            prompt.GetType().GetProperty("Text").SetValue(prompt, text, null);
        }

        /// <summary>Ajoute au sélecteur un scope privé simulé dont les champs reprennent la clé.</summary>
        /// <param name="window">Fenêtre contenant le sélecteur de scope.</param>
        /// <param name="key">Valeur à affecter aux champs d’identité et de libellé.</param>
        /// <returns>Objet de scope ajouté au sélecteur.</returns>
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

        /// <summary>Pompe les messages UI sur le thread STA jusqu’à la fin d’une tâche, avec une limite de cinq secondes.</summary>
        /// <param name="task">Opération asynchrone à attendre.</param>
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
