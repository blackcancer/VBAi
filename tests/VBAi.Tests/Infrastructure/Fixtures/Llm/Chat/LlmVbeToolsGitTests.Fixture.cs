namespace VBAi.Tests.Unit
{
    using System;
    using System.Diagnostics;
    using System.IO;
    using VBAi;

    /// <summary>Hôte VBE simulé contenant les projets associés au test Git.</summary>
    public sealed class ToolGitVbe
    {
        /// <summary>Projets VBA fournis à la session simulée.</summary>
        /// <value>Liste modifiable des projets.</value>
        public System.Collections.Generic.List<ToolGitProject> VBProjects { get; } = new System.Collections.Generic.List<ToolGitProject>();
    }
    /// <summary>Projet VBA de test exposant les propriétés consommées par les opérations Git.</summary>
    public sealed class ToolGitProject
    {
        /// <summary>Nom du projet dans l’hôte.</summary>
        /// <value>Nom utilisé pour identifier le projet.</value>
        public string Name { get; set; }
        /// <summary>Chemin du classeur hôte.</summary>
        /// <value>Chemin utilisé comme scope du dépôt.</value>
        public string FileName { get; set; }
        /// <summary>Mode de fonctionnement du projet.</summary>
        /// <value>Mode par défaut correspondant au mode design.</value>
        public int Mode { get; set; } = 2;
        /// <summary>État de protection du projet.</summary>
        /// <value>Valeur numérique fournie au code de test.</value>
        public int Protection { get; set; }
        /// <summary>Composants VBA du projet simulé.</summary>
        /// <value>Collection de faux composants.</value>
        public global::FakeComponents VBComponents { get; } = new global::FakeComponents();
        /// <summary>Références exposées par le projet.</summary>
        /// <value>Tableau vide par défaut.</value>
        public object[] References { get; } = new object[0];
    }

    /// <summary>Prépare un dépôt temporaire et un projet VBA pour tester les outils Git.</summary>
    public sealed partial class LlmVbeToolsGitTests
    {
        /// <summary>Contient les objets et répertoires isolés utilisés par un test.</summary>
        private sealed class Fixture : IDisposable
        {
            /// <summary>Répertoire temporaire du dépôt, de sa copie locale et du classeur simulé.</summary>
            internal readonly string Root = Path.Combine(Path.GetTempPath(), "CodexToolGit-" + Guid.NewGuid().ToString("N"));
            /// <summary>Dépôt Git local utilisé par les outils.</summary>
            internal readonly MacroGitRepository Repository;
            /// <summary>Adaptateur qui capture les composants du projet VBA.</summary>
            internal readonly VbaGitProject Project;
            /// <summary>Paramètres LLM configurés pour autoriser les modifications.</summary>
            internal readonly LlmSettings Settings = new LlmSettings { VbeEditApproval = "Automatic" };
            /// <summary>Orchestrateur LLM auquel les opérations Git sont injectées.</summary>
            internal readonly LlmVbeTools Tools;
            /// <summary>Projet hôte simulé contenant le module initial.</summary>
            internal readonly global::FakeProject Host;
            /// <summary>Commit initial du dépôt de fixture.</summary>
            internal readonly string Initial;
            /// <summary>Initialise le dépôt temporaire avec un module VBA de base.</summary>
            internal Fixture()
            {
                Directory.CreateDirectory(Root);
                string remote = Path.Combine(Root, "origin.git");
                Git("init", "--bare", remote);
                Repository = new MacroGitRepository(Path.Combine(Root, "cache.git"), "main");
                Repository.Initialize(remote);
                Git("--git-dir=" + Path.Combine(Root, "cache.git"), "config", "user.name", "Coverage Fixture");
                Git("--git-dir=" + Path.Combine(Root, "cache.git"), "config", "user.email", "coverage@example.invalid");
                Host = new global::FakeProject { FileName = Path.Combine(Root, "fixture.xlsm") };
                Host.VBComponents.Add(new global::FakeComponent("Module1", 1, "Attribute VB_Name = \"Module1\"\nOption Explicit\nPublic Const Value = 1\n"));
                Project = new VbaGitProject(() => Host, Host.FileName);
                Initial = Repository.Commit(Project.Capture(), null, "Initial fixture");
                Repository.SetRef(Repository.Head, Initial); Repository.SetRef(MacroGitRepository.Baseline, Initial);
                Tools = new LlmVbeTools(null, null, Settings) { BoundProject = "P", GitOperationsFactory = p => new MacroGitOperations(Project, Repository) };
            }
            /// <summary>Exécute git.exe dans la racine temporaire et renvoie sa sortie standard.</summary>
            /// <param name="arguments">Arguments transmis à Git.</param>
            /// <returns>Sortie standard sans espaces de fin.</returns>
            /// <exception cref="TimeoutException">Git ne termine pas dans les trente secondes.</exception>
            /// <exception cref="InvalidOperationException">Git renvoie un code d’erreur.</exception>
            internal string Git(params string[] arguments)
            {
                var start = new ProcessStartInfo("git.exe", string.Join(" ", Array.ConvertAll(arguments, a => "\"" + a.Replace("\"", "\\\"") + "\"")))
                { WorkingDirectory = Root, UseShellExecute = false, CreateNoWindow = true, RedirectStandardError = true, RedirectStandardOutput = true, RedirectStandardInput = true };
                using (var process = new Process { StartInfo = start })
                {
                    ProcessInput.StartWithoutPreamble(process);
                    process.StandardInput.Close();
                    var output = process.StandardOutput.ReadToEndAsync(); var error = process.StandardError.ReadToEndAsync();
                    if (!process.WaitForExit(30000)) { process.Kill(); throw new TimeoutException("Fixture git timeout"); }
                    if (process.ExitCode != 0) throw new InvalidOperationException(error.GetAwaiter().GetResult());
                    return output.GetAwaiter().GetResult().Trim();
                }
            }
            /// <summary>Vérifie que la racine reste sous le préfixe temporaire avant de supprimer le dépôt.</summary>
            public void Dispose()
            {
                if (!Path.GetFullPath(Root).StartsWith(Path.GetFullPath(Path.GetTempPath()) + "CodexToolGit-", StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("Cleanup escaped fixture root");
                foreach (string file in Directory.GetFiles(Root, "*", SearchOption.AllDirectories)) File.SetAttributes(file, FileAttributes.Normal);
                Directory.Delete(Root, true);
            }
        }
    }
}
