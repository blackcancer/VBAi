namespace CodexVBE.Tests.Infrastructure
{
    using System;
    using System.Threading;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>État du mode VBE renvoyé par l’hôte simulé.</summary>
    public sealed class VbeToolMode
    {
        /// <summary>Mode du projet.</summary>
        /// <value>Code numérique du mode VBE.</value>
        public int Mode { get; set; }
    }
    /// <summary>Informations de signature renvoyées par l’hôte de test.</summary>
    public sealed class VbeToolSignature
    {
        /// <summary>Nom du certificat utilisé.</summary>
        /// <value>Nom du certificat de test.</value>
        public string CertificateName { get; set; }
        /// <summary>Indique que la vérification sans signature a été confirmée.</summary>
        /// <value>Résultat de cette vérification.</value>
        public bool UnsignedVerified { get; set; }
    }
    /// <summary>Résultat de persistance de signature retourné par la fixture.</summary>
    public sealed class VbeToolPersistence
    {
        /// <summary>Indique si les données ont été sauvegardées.</summary>
        /// <value>Valeur de réussite de la persistance.</value>
        public bool Saved { get; set; }
    }

    /// <summary>Données publiques lisibles par le dispatch dynamique de l’assembly de production.</summary>
    public sealed class VbeToolCodeResult
    {
        /// <summary>Contenu stable du module simulé.</summary>
        public string Code { get; set; } = "Option Explicit";
        /// <summary>Version attendue par les arguments de la matrice.</summary>
        public string Sha256 { get; set; } = "value";
        /// <summary>Changements observés par la commande d’historique simulée.</summary>
        public object[] Changes { get; set; } = new object[0];
    }

    /// <summary>Configure les frontières natives VBE et l’exécution hôte pour les tests d’outils.</summary>
    internal static class VbeToolBoundaryFixture
    {
        /// <summary>Remplace les opérations natives par des résultats déterministes de fixture.</summary>
        /// <param name="native">Frontière à configurer.</param>
        internal static void Configure(VbeToolNativeBoundary native)
        {
            native.ReadNavigationSurface = request => new { Available = true };
            native.ChangeNavigationSurface = request => new { Verified = true };
            native.EnsureNoProjectPropertiesDialog = () => { };
            native.ReadProjectProtection = request => new { Available = true };
            native.SetProjectProtection = request => new { CommittedRequested = true };
            native.Capture = stack => new { Native = "capture", Stack = stack };
            native.ReadDebugDialog = () => new { Native = "dialog" };
            native.ChangeDebugItem = request => new { Native = "item", request.Action };
            native.RespondDebugDialog = request => new { Native = "respond", request.Button };
            native.ExecuteImmediate = text => new { Native = "immediate", Text = text };
            native.EnsureNoCompileDialog = () => { };
            native.AwaitCompileDialog = completed => { Assert.IsTrue(completed.Wait(5000)); return null; };
            native.CompleteAddWatch = request => new { Native = "add_watch" };
            native.SelectWatch = request => new { Native = "select_watch" };
            native.CompleteEditWatch = request => new { Native = "edit_watch" };
            native.CompleteQuickWatch = request => new { Native = "quick_watch" };
            native.EnsureNoDebugOptionsDialog = () => { };
            native.ReadVbeOptions = () => new { Native = "vbe_options" };
            native.SetVbeOption = request => new { Native = "set_vbe_option" };
            native.ReadDebugOptions = () => new { Native = "debug_options" };
            native.EnsureNoSignatureDialog = () => { };
            native.ReadSignatureDialog = project => new { Native = "signature_dialog", Project = project };
            native.CompleteProjectSignature = (project, thumbprint, name, unsigned) =>
                new { Native = "signature", Project = project, Thumbprint = thumbprint, CertificateName = name, UnsignedVerified = unsigned };
            native.VerifyWatchRemoved = request => new { Native = "remove_watch" };
        }

        /// <summary>Retourne des réponses déterministes aux commandes de diagnostic et de signature.</summary>
        /// <param name="request">Requête envoyée par un outil VBE.</param>
        /// <returns>Échec pour une requête nulle, sinon succès contenant les données simulées.</returns>
        internal static Response Execute(Request request)
        {
            if (request == null) return Response.Failure("request is null");
            if (request.Command == "debug_state") return Response.Success(new VbeToolMode { Mode = 2 });
            if (request.Command == "sign_project") return Response.Success(new VbeToolSignature { CertificateName = "Disposable", UnsignedVerified = true });
            if (request.Command == "read_module" || request.Command == "native_code_history") return Response.Success(new VbeToolCodeResult());
            if (request.Command == "preview_procedure_rename") return Response.Success(new {
                ExpectedProjectVersion = "value", Edits = new object[0] });
            if (request.Command == "read_project_protection" || request.Command == "set_project_protection")
                return Response.Success(new { ProjectName = request.Project });
            return Response.Success(new { Command = request.Command });
        }
    }
}
