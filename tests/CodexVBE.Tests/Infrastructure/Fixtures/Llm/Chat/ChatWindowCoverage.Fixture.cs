using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using CodexVBE;
using CodexVBE.Tests.Infrastructure;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace CodexVBE.Tests.Unit
{
    /// <summary>Héberge les doubles runtime et les utilitaires visuels partagés par les tests de fenêtre de discussion.</summary>
    public sealed partial class ChatWindowStateTests
    {
        /// <summary>Configure une session VBE simulée et restaure les frontières statiques après le test.</summary>
        private sealed class RuntimeScope : IDisposable
        {
            /// <summary>Délégués statiques de ChatWindow sauvegardés à l’entrée.</summary>
            internal readonly Dictionary<string, object> Defaults = new Dictionary<string, object>();
            /// <summary>Répertoire temporaire de la session.</summary>
            internal readonly string Root = Path.Combine(Path.GetTempPath(), "CodexChat-" + Guid.NewGuid().ToString("N"));
            /// <summary>Paramètres LLM utilisés par la fenêtre simulée.</summary>
            internal readonly LlmSettings Settings = new LlmSettings();
            /// <summary>Gestionnaire remplaçable des commandes hôte.</summary>
            internal Func<Request, Response> Host;
            /// <summary>Session de débogage construite sur le faux VBE.</summary>
            internal readonly VbeSession Session;
            /// <summary>Faux VBE partagé par les tests de cette portée.</summary>
            internal readonly VbeSessionTests.FakeVbe Vbe = new VbeSessionTests.FakeVbe();
            /// <summary>Module VBA de test initial.</summary>
            internal readonly VbeSessionTests.FakeModule Module = new VbeSessionTests.FakeModule("new");
            /// <summary>Transport Codex simulé.</summary>
            internal RuntimeTransport Transport = new RuntimeTransport();
            /// <summary>Nombre de sauvegardes de réglages demandées.</summary>
            internal int Saves;
            /// <summary>Contexte de synchronisation présent avant la portée.</summary>
            private readonly SynchronizationContext originalContext = SynchronizationContext.Current;
            /// <summary>Sauvegarde des paramètres de thème pendant le test.</summary>
            private readonly ThemeScope theme = new ThemeScope();
            /// <summary>Sauvegarde des paramètres de langue pendant le test.</summary>
            private readonly LocalizationScope culture = new LocalizationScope();
            /// <summary>Installe les delegates de test et prépare le faux projet VBA.</summary>
            internal RuntimeScope()
            {
                SynchronizationContext.SetSynchronizationContext(new System.Windows.Threading.DispatcherSynchronizationContext());
                var vbe = Vbe; var project = new VbeSessionTests.FakeProject { Name = "P", FileName = @"C:\Temp\P.xlsm", Mode = 2 }; project.VBComponents.Items.Add(new VbeSessionTests.FakeComponent { Name = "M", Type = 1, CodeModule = Module }); vbe.VBProjects.Add(project); Session = new VbeSession(vbe);
                foreach (var field in typeof(ChatWindow).GetFields(BindingFlags.Static | BindingFlags.NonPublic).Where(x => typeof(Delegate).IsAssignableFrom(x.FieldType))) Defaults[field.Name] = field.GetValue(null);
                ChatWindow.ReadSettings = () => Settings; ChatWindow.WriteSettings = s => Saves++;
                ChatWindow.HistoryPath = () => Path.Combine(Root, "chat.db");
                ChatWindow.ReadModelCatalogue = (p, s) => Task.FromResult(new[] { new LlmModelOption("model", "Model", true, "medium", new[] { new LlmEffortOption("medium", "Medium"), new LlmEffortOption("high", "High") }) });
                ChatWindow.TransportFactory = () => Transport;
                Host = r => Response.Success(r.Command == "list_projects" ? (object)new[] { new { Name = "P", FileName = @"C:\Temp\P.xlsm" } } : r.Command == "debug_state" ? new { SelectedProject = "P", SelectedProjectPath = @"C:\Temp\P.xlsm", ActiveModule = "M", Selection = new { StartLine = 1 } } : r.Command == "read_module" ? new { Code = "Sub A()\nEnd Sub", Sha256 = "sha" } : r.Command == "code_panes" ? (object)new { ActiveCodePane = new { Properties = new { Project = "P", ProjectPath = @"C:\Temp\P.xlsm", Module = "M", Selection = new { StartLine = 1, EndLine = 2, StartColumn = 1, EndColumn = 8 } } } } : new { });
                ChatWindow.ReadHost = (s, r) => Host(r);
                ChatWindow.InvokeTool = (t, n, a) => Task.FromResult(new JavaScriptSerializer().Serialize(Response.Success(new { Compiled = true })));
                ChatWindow.ShowModal = (d, o) => System.Windows.Forms.DialogResult.Cancel;
                ChatWindow.ShowSaveDialog = (d, o) => System.Windows.Forms.DialogResult.Cancel;
                ChatWindow.ShowNotice = (o, t, c, b, i) => System.Windows.Forms.DialogResult.OK;
                ChatWindow.WriteClipboard = s => { };
            }
            /// <summary>Restaure les delegates, le thème et la culture puis supprime le répertoire temporaire.</summary>
            public void Dispose()
            {
                foreach (var entry in Defaults) typeof(ChatWindow).GetField(entry.Key, BindingFlags.Static | BindingFlags.NonPublic).SetValue(null, entry.Value);
                theme.Dispose(); culture.Dispose(); SynchronizationContext.SetSynchronizationContext(originalContext);
                if (Directory.Exists(Root)) Directory.Delete(Root, true);
            }
        }
        /// <summary>Hôte dont l’accès aux projets échoue pour tester l’indisponibilité du catalogue.</summary>
        public sealed class UnavailableReferenceHost
        {
            /// <summary>Provoque une erreur d’E/S à la lecture des projets.</summary>
            /// <value>La lecture lève toujours <see cref="IOException"/>.</value>
            public object VBProjects { get { throw new IOException("reference catalogue unavailable"); } }
        }
        /// <summary>Répond aux appels HTTP de la fenêtre avec un corps configurable.</summary>
        private sealed class RuntimeHttpHandler : System.Net.Http.HttpMessageHandler
        {
            /// <summary>Action appelée avant la production d’une réponse.</summary>
            internal Action BeforeResponse;
            /// <summary>Corps HTTP facultatif de la réponse.</summary>
            internal string Body;
            /// <summary>Indique si la réponse utilise le type de contenu SSE.</summary>
            internal bool Streaming;
            /// <summary>Produit une réponse locale et déclenche le callback de préparation.</summary>
            /// <param name="request">Requête HTTP à traiter.</param>
            /// <param name="cancellationToken">Jeton d’annulation.</param>
            /// <returns>Réponse HTTP avec le corps configuré.</returns>
            protected override Task<System.Net.Http.HttpResponseMessage> SendAsync(System.Net.Http.HttpRequestMessage request, CancellationToken cancellationToken)
            {
                BeforeResponse?.Invoke();
                var content = new System.Net.Http.StringContent(Body ?? "{\"choices\":[{\"message\":{\"role\":\"assistant\",\"content\":\"answer\"}}]}");
                if (Streaming) content.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue("text/event-stream");
                return Task.FromResult(new System.Net.Http.HttpResponseMessage(System.Net.HttpStatusCode.OK) { Content = content });
            }
        }
        /// <summary>Simule les réponses JSON du transport Codex et les événements de tour.</summary>
        private sealed class RuntimeTransport : ICodexAppServerTransport
        {
            /// <summary>Sérialiseur utilisé par le protocole ligne JSON.</summary>
            private readonly JavaScriptSerializer json = new JavaScriptSerializer();
            /// <summary>Événements JSON reçus par le client.</summary>
            public event Action<string> LineReceived;
            /// <summary>Notification de fin de processus simulée.</summary>
            public event Action<Exception> Exited;
            /// <summary>Indique si le transport est démarré.</summary>
            /// <value>État modifié par <see cref="Start"/> et <see cref="Dispose"/>.</value>
            public bool IsRunning { get; private set; }
            /// <summary>Configure les échecs de liste de modèles et de tour, ainsi que la complétion automatique du tour.</summary>
            internal bool FailModels, FailTurn, Complete = true;
            /// <summary>Action appelée avant les événements de fin du tour.</summary>
            internal Action BeforeComplete;
            /// <summary>Modèles renvoyés par la requête de liste.</summary>
            internal object[] Models = { new { model = "model", displayName = "Model", isDefault = true, defaultReasoningEffort = "medium", supportedReasoningEfforts = new[] { new { reasoningEffort = "medium", description = "Medium" } } } };
            /// <summary>Marque le transport comme actif.</summary>
            public void Start() { IsRunning = true; }
            /// <summary>Marque le transport comme arrêté.</summary>
            public void Dispose() { IsRunning = false; }
            /// <summary>Sérialise une valeur et la transmet comme ligne reçue.</summary>
            /// <param name="value">Objet de réponse ou d’événement à émettre.</param>
            internal void Emit(object value) { LineReceived?.Invoke(json.Serialize(value)); }
            /// <summary>Signale la fin du transport par une erreur d’E/S.</summary>
            internal void Exit() { Exited?.Invoke(new IOException("closed")); }
            /// <summary>Traite une requête JSON et émet le résultat et les événements simulés associés.</summary>
            /// <param name="line">Ligne JSON envoyée par le client.</param>
            public void Send(string line)
            {
                var msg = (IDictionary<string, object>)json.DeserializeObject(line);
                if (!msg.ContainsKey("id") || !msg.ContainsKey("method")) return;
                var id = msg["id"]; var method = Convert.ToString(msg["method"]);
                if (method == "model/list" && FailModels) { Emit(new { id, error = new { message = "models failed" } }); return; }
                object result = method == "account/read" ? (object)new { account = new { type = "chatgpt" } } : method == "model/list" ? new { data = Models, nextCursor = (string)null } : method == "thread/start" || method == "thread/resume" ? (object)new { thread = new { id = "thread" } } : method == "turn/start" ? (object)new { turn = new { id = "turn" } } : new { };
                if (method == "turn/start") Emit(new { method = "turn/started", @params = new { threadId = "thread", turn = new { id = "turn" } } });
                Emit(new { id, result });
                if (method == "turn/start" && Complete) { BeforeComplete?.Invoke(); Emit(new { method = "item/agentMessage/delta", @params = new { threadId = "thread", itemId = "answer", delta = "answer" } }); Emit(new { method = "item/completed", @params = new { threadId = "thread", item = new { type = "agentMessage", phase = "final", id = "answer", text = "answer" } } }); Emit(new { method = "turn/completed", @params = new { threadId = "thread", turn = new { status = FailTurn ? "failed" : "completed", error = FailTurn ? new { message = "turn failed" } : null } } }); }
                if (method == "turn/interrupt") Emit(new { method = "turn/completed", @params = new { threadId = "thread", turn = new { status = "interrupted", error = (object)null } } });
            }
        }
        /// <summary>Recherche récursivement le premier élément WPF du type demandé.</summary>
        /// <typeparam name="T">Type d’élément visuel à rechercher.</typeparam>
        /// <param name="root">Racine de l’arbre visuel.</param>
        /// <returns>Premier descendant correspondant, ou <see langword="null"/>.</returns>
        private static T Visual<T>(DependencyObject root) where T : DependencyObject { if (root is T value) return value; for (int i = 0; i < System.Windows.Media.VisualTreeHelper.GetChildrenCount(root); i++) { var result = Visual<T>(System.Windows.Media.VisualTreeHelper.GetChild(root, i)); if (result != null) return result; } return null; }
        /// <summary>Déclenche un événement ScrollChanged synthétique sur un élément WPF.</summary>
        /// <param name="window">Fenêtre associée à l’événement.</param><param name="source">Élément destinataire.</param><param name="vertical">Position verticale simulée.</param><param name="extent">Étendue verticale simulée.</param>
        private static void RaiseScroll(ChatWindow window, UIElement source, double vertical, double extent) { var args = (ScrollChangedEventArgs)Activator.CreateInstance(typeof(ScrollChangedEventArgs), BindingFlags.Instance | BindingFlags.NonPublic, null, new object[] { new Vector(0, vertical), new Vector(0, vertical), new Size(100, 100), new Vector(0, extent), new Size(100, 30), new Vector(0, 0) }, null); args.RoutedEvent = ScrollViewer.ScrollChangedEvent; source.RaiseEvent(args); }
        /// <summary>Déclenche le tick d’un DispatcherTimer par réflexion.</summary>
        /// <param name="timer">Timer dont le callback doit s’exécuter.</param>
        private static void TimerTick(System.Windows.Threading.DispatcherTimer timer) { var method = typeof(System.Windows.Threading.DispatcherTimer).GetMethod("FireTick", BindingFlags.Instance | BindingFlags.NonPublic); method.Invoke(timer, new object[method.GetParameters().Length]); }
        /// <summary>Simule un clic WinForms en appelant le gestionnaire protégé du contrôle.</summary>
        /// <param name="control">Contrôle ciblé.</param>
        private static void Click(System.Windows.Forms.Control control) { typeof(System.Windows.Forms.Control).GetMethod("OnClick", BindingFlags.Instance | BindingFlags.NonPublic).Invoke(control, new object[] { EventArgs.Empty }); }
        /// <summary>Énumère récursivement un élément WPF puis ses descendants logiques.</summary>
        /// <param name="element">Élément racine.</param>
        /// <returns>Éléments du sous-arbre dans l’ordre de parcours.</returns>
        private static IEnumerable<FrameworkElement> Descendants(FrameworkElement element)
        {
            yield return element;
            foreach (var child in LogicalTreeHelper.GetChildren(element).OfType<FrameworkElement>()) foreach (var nested in Descendants(child)) yield return nested;
        }
        /// <summary>Simule un clic sur un bouton WPF.</summary>
        /// <param name="button">Bouton qui reçoit l’événement routé.</param>
        private static void WpfClick(System.Windows.Controls.Button button) { button.RaiseEvent(new RoutedEventArgs(System.Windows.Controls.Button.ClickEvent)); }
        /// <summary>Exécute le gestionnaire de touche de prompt avec une source WPF temporaire.</summary>
        /// <param name="window">Fenêtre qui reçoit la touche.</param>
        /// <param name="key">Touche à transmettre au gestionnaire.</param>
        /// <returns>Arguments après traitement, dont l’état Handled peut être vérifié.</returns>
        private static KeyEventArgs RunKey(ChatWindow window, Key key)
        {
            var prompt = Get<System.Windows.Controls.TextBox>(window, "prompt");
            using (var source = new System.Windows.Interop.HwndSource(new System.Windows.Interop.HwndSourceParameters("test") { Width = 1, Height = 1, PositionX = -10000, PositionY = -10000 }))
            {
                // Source lifetime only supplies keyboard event routing; no input is synthesized.
                var args = new KeyEventArgs(Keyboard.PrimaryDevice, source, 0, key) { RoutedEvent = Keyboard.PreviewKeyDownEvent }; Call(window, "PromptKeyDown", prompt, args); return args;
            }
        }
    }
}
