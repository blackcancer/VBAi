namespace VBAi.Tests.Unit
{
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading;
    using System.Web.Script.Serialization;
    using VBAi;
    using VBAi.Tests.Infrastructure;

    /// <summary>Fournit un proxy autour des vraies orchestrations d’outils et des frontières hôte simulées.</summary>
    public sealed partial class LlmVbeToolsBoundaryTests
    {
        /// <summary>Sérialiseur JSON configuré pour les schémas volumineux des outils.</summary>
        private static readonly JavaScriptSerializer Json = new JavaScriptSerializer { MaxJsonLength = 10 * 1024 * 1024 };
        /// <summary>Convertit une valeur JSON en dictionnaire.</summary>
        /// <param name="value">Valeur désérialisée.</param>
        /// <returns>Dictionnaire des propriétés.</returns>
        private static IDictionary<string, object> Dict(object value) { return (IDictionary<string, object>)value; }
        /// <summary>Configure les outils et remplace uniquement les appels aux frontières natives.</summary>
        private sealed class ToolFixture
        {
            /// <summary>Paramètres de test avec approbation automatique des modifications VBA.</summary>
            internal readonly LlmSettings Settings = new LlmSettings { VbeEditApproval = "Automatic" };
            /// <summary>Orchestrateur d’outils dont le comportement métier est conservé.</summary>
            internal readonly LlmVbeTools Tools;
            /// <summary>Crée les outils sur un faux VBE et configure leurs délégués natifs.</summary>
            internal ToolFixture()
            {
                Tools = new LlmVbeTools(new VbeSession(new VbeSessionTests.FakeVbe()), null, Settings);
                Tools.ImmediateOwnerDispatch = action => action();
                VbeToolBoundaryFixture.Configure(Tools.Native);
                Tools.Execute = VbeToolBoundaryFixture.Execute;
                Tools.CaptureSignaturePersistence = p => () => { };
                Tools.PersistSignature = (p, authorize) => { authorize(); return new VBAi.Tests.Infrastructure.VbeToolPersistence { Saved = true }; };
            }
        }
        // Each tool retains its real orchestrator; only its host/native boundaries are replaced.
        /// <summary>Crée le proxy d’accès aux outils pour les assertions de test.</summary>
        /// <returns>Proxy configuré sur une instance d’outils.</returns>
        private static ToolProxy Create() { return new ToolProxy(); }
        /// <summary>Expose les frontières et opérations injectables sans remplacer les orchestrations réelles.</summary>
        private sealed class ToolProxy
        {
            /// <summary>Fixture qui possède les outils et leurs paramètres.</summary>
            private readonly ToolFixture fixture = new ToolFixture();
            /// <summary>Paramètres utilisés par les outils.</summary>
            /// <value>Paramètres détenus par la fixture.</value>
            internal LlmSettings Settings { get { return fixture.Settings; } }
            /// <summary>Frontières natives remplaçables des outils.</summary>
            /// <value>Instance injectée dans les orchestrations.</value>
            internal VbeToolNativeBoundary Native { get { return fixture.Tools.Native; } }
            /// <summary>Définit le délégué d’exécution des requêtes hôte.</summary>
            /// <value>Délégué affecté à l’orchestrateur.</value>
            internal Func<Request, Response> Execute { set { fixture.Tools.Execute = value; } }
            /// <summary>Définit le délégué de persistance de signature.</summary>
            /// <value>Fonction appelée pour persister une signature.</value>
            internal Func<string, Action, object> PersistSignature { set { fixture.Tools.PersistSignature = value; } }
            /// <summary>Définit le nom du fournisseur courant.</summary>
            /// <value>Nom du fournisseur défini sur l’orchestrateur.</value>
            internal string CurrentProviderName { set { fixture.Tools.CurrentProviderName = value; } }
            /// <summary>Définit le projet lié à la conversation.</summary>
            /// <value>Nom du projet lié aux opérations.</value>
            internal string BoundProject { set { fixture.Tools.BoundProject = value; } }
            /// <summary>Définit le mode courant de la conversation.</summary>
            /// <value>Mode transmis à l’orchestrateur.</value>
            internal ChatMode Mode { set { fixture.Tools.Mode = value; } }
            /// <summary>Définit le gestionnaire de confirmation de fichier.</summary>
            /// <value>Délégué appelé pour confirmer une opération sur fichier.</value>
            internal Func<System.Windows.Forms.IWin32Window, string, string, System.Windows.Forms.DialogResult> ConfirmFile { set { fixture.Tools.ConfirmFile = value; } }
            /// <summary>Définit le gestionnaire d’approbation des actions VBE.</summary>
            /// <value>Délégué appelé pour afficher l’approbation VBE.</value>
            internal Func<VbeApprovalDialog, System.Windows.Forms.IWin32Window, System.Windows.Forms.DialogResult> ShowApproval { set { fixture.Tools.ShowApproval = value; } }
            /// <summary>Enregistre une demande de l’utilisateur dans le contexte des outils.</summary>
            /// <param name="text">Texte de la demande utilisateur.</param>
            internal void NoteUserRequest(string text) { fixture.Tools.NoteUserRequest(text); }
            /// <summary>Invoque un outil de façon synchrone.</summary>
            /// <param name="name">Nom de l’outil.</param>
            /// <param name="arguments">Arguments JSON de l’outil.</param>
            /// <returns>Réponse sérialisée de l’outil.</returns>
            internal string Invoke(string name, string arguments) { return fixture.Tools.Invoke(name, arguments); }
            /// <summary>Invoque un outil de façon asynchrone.</summary>
            /// <param name="name">Nom de l’outil.</param>
            /// <param name="arguments">Arguments JSON de l’outil.</param>
            /// <returns>Tâche qui produit la réponse sérialisée.</returns>
            internal System.Threading.Tasks.Task<string> InvokeAsync(string name, string arguments) { return fixture.Tools.InvokeAsync(name, arguments); }
        }
        /// <summary>Construit des arguments JSON valides à partir du schéma d’un outil nommé.</summary>
        /// <param name="name">Nom de l’outil dont le schéma définit les champs.</param>
        /// <returns>Dictionnaire d’arguments par défaut adaptés aux types du schéma.</returns>
        private static Dictionary<string, object> Arguments(string name)
        {
            var definition = LlmVbeTools.Definitions.Select(d => Dict(Dict(Json.DeserializeObject(Json.Serialize(d)))["function"]))
                .Single(d => (string)d["name"] == name);
            var fields = Dict(Dict(definition["parameters"])["properties"]);
            return fields.ToDictionary(f => f.Key, f => f.Key == "Value" ? (object)"value" :
                f.Key == "Arguments" ? new object[0] : f.Key == "ArgumentNames" ? new string[0] : f.Key == "Items" ? new string[0] : f.Key == "Rows" ? (object)new string[0][] : f.Key == "PathSegments" ? new[] { "item" } :
                (string)Dict(f.Value)["type"] == "integer" ? (object)2 :
                (string)Dict(f.Value)["type"] == "number" ? (object)1.5 :
                (string)Dict(f.Value)["type"] == "boolean" ? (object)true :
                f.Key == "Path" ? @"C:\Temp\fixture.bas" : f.Key == "Project" ? "P" : "value");
        }
        /// <summary>Construit des arguments adaptés aux outils dont les paramètres asynchrones ont des champs spécifiques.</summary>
        /// <param name="name">Nom de l’outil.</param>
        /// <returns>Arguments spécifiques, ou ceux générés par <see cref="Arguments"/>.</returns>
        private static Dictionary<string, object> AsyncArguments(string name)
        {
            if (name == "debug_item") return new Dictionary<string, object> { ["Pane"] = "locals", ["Action"] = "expand", ["PathSegments"] = new[] { "item" } };
            if (name == "immediate_execute") return new Dictionary<string, object> { ["Project"] = "P", ["ExpectedMode"] = 2, ["Text"] = "Debug.Print 1" };
            if (name == "debug_dialog") return new Dictionary<string, object>();
            if (name == "respond_debug_dialog") return new Dictionary<string, object> { ["Diagnostic"] = "fixture", ["Button"] = "ok" };
            return Arguments(name);
        }
        /// <summary>Vérifie qu’une réponse d’outil sérialisée indique une réussite.</summary>
        /// <param name="json">Réponse JSON à examiner.</param>
        /// <param name="context">Contexte ajouté au message d’assertion.</param>
        private static void Success(string json, string context)
        { var response = Json.Deserialize<Response>(json); Assert.IsTrue(response.Ok, context + ": " + response.Error); }
        /// <summary>Vérifie qu’une réponse d’outil contient un échec avec un message non vide.</summary>
        /// <param name="json">Réponse JSON à examiner.</param>
        /// <param name="context">Contexte ajouté au message d’assertion.</param>
        private static void Failed(string json, string context)
        { var response = Json.Deserialize<Response>(json); Assert.IsFalse(response.Ok, context); Assert.IsFalse(string.IsNullOrWhiteSpace(response.Error), context); }
        /// <summary>Valide une réponse positive et renvoie son objet de données.</summary>
        /// <param name="json">Réponse JSON d’un outil.</param>
        /// <returns>Dictionnaire contenu dans la propriété Data.</returns>
        private static IDictionary<string, object> Data(string json)
        { Success(json, "response"); return Dict(Json.Deserialize<Response>(json).Data); }
        /// <summary>Exécute immédiatement les callbacks postés sur le contexte.</summary>
        private sealed class ImmediateContext : SynchronizationContext
        {
            /// <summary>Exécute le callback sur le thread appelant.</summary>
            /// <param name="callback">Délégué à exécuter.</param>
            /// <param name="state">État transmis au callback.</param>
            public override void Post(SendOrPostCallback callback, object state) { callback(state); }
        }
        /// <summary>Diffère le premier callback posté et envoie les suivants sur le pool de threads.</summary>
        private sealed class DeferredContext : SynchronizationContext
        {
            /// <summary>Nombre de callbacks soumis au contexte.</summary>
            private int posts;
            private Action pending;
            internal void ReleasePending() { pending?.Invoke(); pending = null; }
            /// <summary>Ignore le premier callback puis programme les suivants sur le pool de threads.</summary>
            /// <param name="callback">Délégué à exécuter.</param>
            /// <param name="state">État transmis au callback.</param>
            public override void Post(SendOrPostCallback callback, object state)
            {
                if (Interlocked.Increment(ref posts) == 1) pending = () => callback(state);
                else ThreadPool.QueueUserWorkItem(_ => callback(state));
            }
        }
    }
}
