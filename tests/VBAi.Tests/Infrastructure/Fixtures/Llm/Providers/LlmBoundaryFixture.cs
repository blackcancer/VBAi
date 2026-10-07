using Microsoft.CSharp;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.CodeDom.Compiler;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows.Forms;

namespace VBAi.Tests.Infrastructure
{
    /// <summary>Installe des fournisseurs et réglages isolés, puis restaure l’état global après les tests LLM.</summary>
    internal sealed class LlmBoundaryScope : IDisposable
    {
        /// <summary>Délégués statiques sauvegardés pour restauration.</summary>
        private readonly Dictionary<FieldInfo, object> defaults = new Dictionary<FieldInfo, object>();
        /// <summary>Valeurs initiales des variables d’environnement modifiées.</summary>
        private readonly Dictionary<string, string> environment = new Dictionary<string, string>();
        /// <summary>Portée de sauvegarde des préférences de thème.</summary>
        private readonly ThemeScope theme = new ThemeScope();
        /// <summary>Portée de sauvegarde de la culture.</summary>
        private readonly LocalizationScope culture = new LocalizationScope();
        /// <summary>Contexte de synchronisation présent avant l’installation de la fixture.</summary>
        private readonly SynchronizationContext context = SynchronizationContext.Current;
        /// <summary>Répertoire temporaire des exécutables et marqueurs du test.</summary>
        internal readonly string Root = Path.Combine(Path.GetTempPath(), "CodexLlm-" + Guid.NewGuid().ToString("N"));
        /// <summary>Textes passés aux fenêtres de notification simulées.</summary>
        internal readonly List<string> Notices = new List<string>();
        /// <summary>Fournisseurs dont l’action de connexion a été demandée.</summary>
        internal readonly List<string> Logins = new List<string>();
        /// <summary>Nombre d’enregistrements de paramètres demandés.</summary>
        internal int Saves;
        /// <summary>Crée un répertoire temporaire et remplace les frontières LLM globales par des doublures.</summary>
        internal LlmBoundaryScope()
        {
            Directory.CreateDirectory(Root);
            foreach (var field in typeof(LlmSettingsWindow).GetFields(BindingFlags.Static | BindingFlags.NonPublic).Where(x => typeof(Delegate).IsAssignableFrom(x.FieldType))) defaults[field] = field.GetValue(null);
            var handler = typeof(LlmChatClient).GetField("HttpHandlerFactory", BindingFlags.Static | BindingFlags.NonPublic); defaults[handler] = handler.GetValue(null);
            foreach (var name in LlmProvider.All.SelectMany(p => new[] { p.KeyVariable, p.ModelVariable, p.EndpointVariable }).Concat(new[] { "AZURE_OPENAI_ENTRA_TOKEN", "VBAi_COPILOT_CLI", "VBAi_CODEX_CLI", "VBAi_TEST_COPILOT_MODE", "VBAi_TEST_COPILOT_MARKER" }).Where(x => x != null).Distinct()) { environment[name] = Environment.GetEnvironmentVariable(name); Environment.SetEnvironmentVariable(name, null); }
            LlmSettingsWindow.WriteSettings = s => Saves++;
            LlmSettingsWindow.StartCopilotLogin = () => Logins.Add("copilot"); LlmSettingsWindow.StartCodexLogin = () => Logins.Add("codex");
            LlmSettingsWindow.ReadCopilotStatus = () => Task.FromResult("Copilot fixture connected");
            LlmSettingsWindow.ReadCodexStatus = () => Task.FromResult(new CodexAccountStatus(true, "ChatGPT fixture connected"));
            LlmSettingsWindow.ShowNotice = (o, t, c, b, i) => { Notices.Add(t); return DialogResult.OK; };
            LlmChatClient.HttpHandlerFactory = () => new LlmHttpFixture("{\"data\":[]}");
        }
        /// <summary>Crée une fenêtre de paramètres hors écran et remplace son service GitHub.</summary>
        /// <param name="settings">Paramètres de fournisseur affichés dans la fenêtre.</param>
        /// <returns>Fenêtre configurée pour le test.</returns>
        internal LlmSettingsWindow Window(LlmSettings settings)
        {
            if (Thread.CurrentThread.GetApartmentState() == ApartmentState.STA) SynchronizationContext.SetSynchronizationContext(new System.Windows.Threading.DispatcherSynchronizationContext());
            var window = new LlmSettingsWindow(settings) { Left = -10000, Top = -10000, ShowInTaskbar = false };
            Set(window, "githubService", new GitHubAccountService((command, token) => Task.FromResult("")));
            return window;
        }
        /// <summary>Compile et sélectionne le processus Copilot de fixture avec le mode demandé.</summary>
        /// <param name="mode">Scénario protocolaire activé par la variable d’environnement du processus.</param>
        internal void UseCopilot(string mode = "normal")
        {
            var executable = Path.Combine(Root, "copilot-fixture.exe");
            if (!File.Exists(executable))
                using (var compiler = new CSharpCodeProvider())
                {
                    var options = new CompilerParameters(new[] { "System.dll", "System.Core.dll", "System.Web.Extensions.dll" }, executable) { GenerateExecutable = true, GenerateInMemory = false, CompilerOptions = "/target:winexe /optimize+" };
                    var result = compiler.CompileAssemblyFromSource(options, CopilotFixtureProgram.Source);
                    Assert.IsFalse(result.Errors.HasErrors, string.Join("\n", result.Errors.Cast<CompilerError>().Select(x => x.ToString())));
                }
            Environment.SetEnvironmentVariable("VBAi_COPILOT_CLI", executable); Environment.SetEnvironmentVariable("VBAi_TEST_COPILOT_MODE", mode); Environment.SetEnvironmentVariable("VBAi_TEST_COPILOT_MARKER", Path.Combine(Root, "login.marker"));
        }
        /// <summary>Retrouve un fournisseur par son nom affiché.</summary>
        /// <param name="name">Nom exact du fournisseur.</param>
        /// <returns>Fournisseur correspondant.</returns>
        internal static LlmProvider Provider(string name) { return LlmProvider.All.Single(x => x.Name == name); }
        /// <summary>Appelle la méthode d’instance non publique portant le nom et l’arité indiqués.</summary>
        /// <param name="target">Objet cible.</param>
        /// <param name="name">Nom de la méthode.</param>
        /// <param name="args">Arguments transmis à l’appel.</param>
        /// <returns>Valeur renvoyée par la méthode.</returns>
        internal static object Call(object target, string name, params object[] args) { return target.GetType().GetMethods(BindingFlags.NonPublic | BindingFlags.Public | BindingFlags.Instance).Single(x => x.Name == name && x.GetParameters().Length == args.Length).Invoke(target, args); }
        /// <summary>Lit un champ privé d’un objet.</summary>
        /// <typeparam name="T">Type attendu du champ.</typeparam>
        /// <param name="target">Objet contenant le champ.</param>
        /// <param name="name">Nom du champ.</param>
        /// <returns>Valeur du champ convertie en <typeparamref name="T"/>.</returns>
        internal static T Get<T>(object target, string name) { return (T)target.GetType().GetField(name, BindingFlags.Instance | BindingFlags.NonPublic).GetValue(target); }
        /// <summary>Écrit un champ privé d’un objet.</summary>
        /// <param name="target">Objet contenant le champ.</param>
        /// <param name="name">Nom du champ.</param>
        /// <param name="value">Valeur à affecter.</param>
        internal static void Set(object target, string name, object value) { target.GetType().GetField(name, BindingFlags.Instance | BindingFlags.NonPublic).SetValue(target, value); }
        /// <summary>Déclenche le gestionnaire de clic protégé d’un contrôle WinForms.</summary>
        /// <param name="button">Contrôle dont le clic doit être simulé.</param>
        internal static void Click(Control button) { typeof(Control).GetMethod("OnClick", BindingFlags.Instance | BindingFlags.NonPublic).Invoke(button, new object[] { EventArgs.Empty }); }
        /// <summary>Pompe la boucle WinForms jusqu’à la fin de la tâche ou l’expiration du délai.</summary>
        /// <param name="task">Tâche à attendre.</param>
        /// <exception cref="AssertFailedException">La tâche n’est pas terminée dans les douze secondes.</exception>
        internal static void Pump(Task task)
        {
            var deadline = DateTime.UtcNow.AddSeconds(12);
            while (!task.IsCompleted && DateTime.UtcNow < deadline) { Application.DoEvents(); Thread.Sleep(1); }
            Assert.IsTrue(task.IsCompleted, "The isolated provider operation timed out."); task.GetAwaiter().GetResult();
        }
        /// <summary>Convertit une valeur en dictionnaire JSON via une sérialisation puis désérialisation.</summary>
        /// <param name="data">Objet à convertir.</param>
        /// <returns>Dictionnaire de l’objet JSON désérialisé.</returns>
        internal static IDictionary<string, object> Object(object data) { return new JavaScriptSerializer().DeserializeObject(new JavaScriptSerializer().Serialize(data)) as IDictionary<string, object>; }
        /// <summary>Restaure les délégués, variables d’environnement et contextes puis supprime le répertoire temporaire.</summary>
        public void Dispose()
        {
            foreach (var pair in defaults) pair.Key.SetValue(null, pair.Value);
            foreach (var pair in environment) Environment.SetEnvironmentVariable(pair.Key, pair.Value);
            theme.Dispose(); culture.Dispose(); SynchronizationContext.SetSynchronizationContext(context);
            // Process.Kill returns before Windows releases the executable image.
            for (int attempt = 0; Directory.Exists(Root); attempt++)
            {
                try { Directory.Delete(Root, true); }
                catch (IOException) { if (attempt >= 100) throw; Thread.Sleep(10); }
                catch (UnauthorizedAccessException) { if (attempt >= 100) throw; Thread.Sleep(10); }
            }
        }
    }

    /// <summary>Répond aux appels HTTP LLM à partir d’une file de réponses contrôlée par le test.</summary>
    internal sealed class LlmHttpFixture : HttpMessageHandler
    {
        /// <summary>Décrit le contenu, le type MIME et le statut d’une réponse de fixture.</summary>
        internal sealed class Reply
        {
            /// <summary>Corps retourné et type MIME attribué au contenu de la réponse.</summary>
            internal string Body, MediaType = "application/json";
            /// <summary>Indique si l’en-tête Content-Type doit être omis.</summary>
            internal bool OmitContentType;
            /// <summary>Code HTTP de la réponse.</summary>
            internal HttpStatusCode Status = HttpStatusCode.OK;
            /// <summary>Crée une réponse avec le corps spécifié.</summary>
            /// <param name="body">Contenu textuel de la réponse.</param>
            internal Reply(string body) { Body = body; }
        }
        /// <summary>Réponses à consommer dans l’ordre des requêtes.</summary>
        internal readonly Queue<Reply> Replies = new Queue<Reply>();
        /// <summary>URI de chaque requête reçue.</summary>
        internal readonly List<Uri> Uris = new List<Uri>();
        /// <summary>Corps de chaque requête reçue.</summary>
        internal readonly List<string> Bodies = new List<string>();
        /// <summary>En-têtes de chaque requête reçue.</summary>
        internal readonly List<Dictionary<string, string>> Headers = new List<Dictionary<string, string>>();
        /// <summary>Action appelée juste avant la production de la réponse.</summary>
        internal Action BeforeResponse;
        /// <summary>Prépare les réponses HTTP dans l’ordre fourni.</summary>
        /// <param name="bodies">Corps JSON ou texte à renvoyer.</param>
        internal LlmHttpFixture(params string[] bodies) { foreach (var body in bodies) Replies.Enqueue(new Reply(body)); }
        /// <summary>Enregistre la requête et construit la prochaine réponse préconfigurée.</summary>
        /// <param name="request">Requête HTTP à traiter.</param>
        /// <param name="token">Jeton d’annulation de la requête.</param>
        /// <returns>Réponse HTTP décrite par le prochain objet <see cref="Reply"/>.</returns>
        /// <exception cref="InvalidOperationException">Aucune réponse n’est disponible dans la file.</exception>
        protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken token)
        {
            Uris.Add(request.RequestUri); Bodies.Add(request.Content == null ? "" : await request.Content.ReadAsStringAsync()); Headers.Add(request.Headers.ToDictionary(x => x.Key, x => string.Join(",", x.Value), StringComparer.OrdinalIgnoreCase)); BeforeResponse?.Invoke();
            if (Replies.Count == 0) throw new InvalidOperationException("No fixture response remains.");
            var reply = Replies.Dequeue(); var content = new StringContent(reply.Body, System.Text.Encoding.UTF8, reply.MediaType); if (reply.OmitContentType) content.Headers.ContentType = null; return new HttpResponseMessage(reply.Status) { Content = content };
        }
    }
    /// <summary>Contexte de synchronisation qui diffère les callbacks jusqu’à un appel explicite à <see cref="Drain"/>.</summary>
    internal sealed class LlmQueuedContext : SynchronizationContext
    {
        /// <summary>Callbacks en attente d’exécution.</summary>
        private readonly Queue<Action> work = new Queue<Action>();
        /// <summary>Ajoute un callback à la file sans l’exécuter immédiatement.</summary>
        /// <param name="callback">Callback à exécuter lors du vidage.</param>
        /// <param name="state">État transmis au callback.</param>
        public override void Post(SendOrPostCallback callback, object state) { lock (work) work.Enqueue(() => callback(state)); }
        /// <summary>Nombre de callbacks en attente.</summary>
        /// <value>Nombre d’actions actuellement dans la file.</value>
        internal int Count { get { lock (work) return work.Count; } }
        /// <summary>Exécute les callbacks jusqu’à ce que la file soit vide.</summary>
        internal void Drain() { while (true) { Action action; lock (work) { if (work.Count == 0) return; action = work.Dequeue(); } action(); } }
    }
}
