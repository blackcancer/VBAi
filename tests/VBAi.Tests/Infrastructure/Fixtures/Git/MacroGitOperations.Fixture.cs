namespace VBAi.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.IO;
    using System.Text;
    using VBAi;

    /// <summary>Fournit un dépôt Git temporaire et un projet VBA simulé aux tests d’opérations.</summary>
    public sealed partial class MacroGitOperationsTests
    {
        /// <summary>Crée et détruit l’environnement Git et le projet simulé pour un test.</summary>
        internal sealed class Fixture : IDisposable
        {
            /// <summary>Répertoire temporaire réservé au dépôt et au dépôt distant du test.</summary>
            internal readonly string Root = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "git-branches", Guid.NewGuid().ToString("N"));
            /// <summary>Chemins du dépôt local et du dépôt distant temporaire de la fixture.</summary>
            internal readonly string Cache, Remote;
            /// <summary>Projet VBA simulé par des composants en mémoire.</summary>
            internal readonly global::FakeProject Host;
            /// <summary>Adaptateur reliant le projet simulé au dépôt.</summary>
            internal readonly VbaGitProject Project;
            /// <summary>Dépôt local initialisé sur la branche principale.</summary>
            internal readonly MacroGitRepository Repository;
            /// <summary>Service testé pour les opérations de synchronisation VBA et Git.</summary>
            internal readonly MacroGitOperations Operations;
            /// <summary>Initialise le dépôt distant, le dépôt local et un projet contenant un module de base.</summary>
            internal Fixture()
            {
                Directory.CreateDirectory(Root);
                Cache = Path.Combine(Root, "cache"); Remote = Path.Combine(Root, "origin.git");
                Git("init", "--bare", Remote);
                Repository = new MacroGitRepository(Cache, "main"); Repository.Initialize(Remote);
                Git("--git-dir=" + Cache, "config", "user.name", "Coverage Fixture");
                Git("--git-dir=" + Cache, "config", "user.email", "coverage@example.invalid");
                Host = new global::FakeProject { FileName = Path.Combine(Root, "fixture.xlsm") };
                Host.VBComponents.Add(new global::FakeComponent("Module1", 1, "Attribute VB_Name = \"Module1\"\nOption Explicit\nPublic Const Value = 1\n"));
                Project = new VbaGitProject(() => Host, Host.FileName);
                Operations = new MacroGitOperations(Project, Repository);
            }
            /// <summary>Exécute Git dans le répertoire temporaire et échoue si la commande dépasse le délai ou retourne une erreur.</summary>
            /// <param name="arguments">Arguments passés à git.exe.</param>
            /// <exception cref="TimeoutException">La commande ne termine pas dans les trente secondes.</exception>
            /// <exception cref="InvalidOperationException">Git retourne un code de sortie non nul.</exception>
            internal void Git(params string[] arguments)
            {
                var info = new ProcessStartInfo("git.exe", string.Join(" ", Array.ConvertAll(arguments, a => "\"" + a.Replace("\"", "\\\"") + "\"")))
                { WorkingDirectory = Root, UseShellExecute = false, CreateNoWindow = true, RedirectStandardError = true, RedirectStandardOutput = true };
                using (var process = Process.Start(info))
                {
                    var output = process.StandardOutput.ReadToEndAsync(); var error = process.StandardError.ReadToEndAsync();
                    if (!process.WaitForExit(30000)) { process.Kill(); throw new TimeoutException("Fixture Git timed out"); }
                    if (process.ExitCode != 0) throw new InvalidOperationException(error.GetAwaiter().GetResult());
                    output.GetAwaiter().GetResult();
                }
            }
            /// <summary>Capture le projet et remplace la constante du module pour produire un instantané distinct.</summary>
            /// <param name="value">Valeur de constante VBA à écrire dans l’instantané.</param>
            /// <returns>Instantané construit à partir des fichiers du projet.</returns>
            internal VbaGitSnapshot Snapshot(string value)
            {
                var files = Project.Capture().Serialize();
                files["Module1.bas"] = Encoding.UTF8.GetBytes("Attribute VB_Name = \"Module1\"\nOption Explicit\nPublic Const Value = " + value + "\n");
                return VbaGitSnapshot.Read(files);
            }
            /// <summary>Crée un commit de fixture dans le dépôt local.</summary>
            /// <param name="snapshot">Instantané VBA à enregistrer.</param>
            /// <param name="parent">Commit parent facultatif.</param>
            /// <returns>Identifiant du commit créé.</returns>
            internal string Commit(VbaGitSnapshot snapshot, string parent = null)
            { return Repository.Commit(snapshot, parent, "Fixture commit"); }
            /// <summary>Enregistre l’état courant comme HEAD et comme base de comparaison.</summary>
            /// <returns>Identifiant du commit initial créé.</returns>
            internal string Seed()
            {
                string commit = Commit(Project.Capture());
                Repository.SetRef(Repository.Head, commit); Repository.SetRef(MacroGitRepository.Baseline, commit);
                return commit;
            }
            /// <summary>Exécute ImportAsync par réflexion et propage l’exception de tâche d’origine.</summary>
            /// <param name="target">Instantané à importer.</param>
            /// <param name="expected">Instantané attendu avant mutation.</param>
            /// <param name="rollback">Indique si l’import est un retour arrière.</param>
            /// <returns><see langword="null"/> après réussite; une erreur est propagée en cas d’échec.</returns>
            internal object Import(VbaGitSnapshot target, VbaGitSnapshot expected, bool rollback = false)
            {
                try
                {
                    var task = (System.Threading.Tasks.Task)typeof(MacroGitOperations).GetMethod("ImportAsync", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)
                        .Invoke(Operations, new object[] { target, expected, rollback });
                    task.GetAwaiter().GetResult(); return null;
                }
                catch (System.Reflection.TargetInvocationException ex)
                { System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(ex.InnerException).Throw(); throw; }
            }
            /// <summary>Libère les opérations puis supprime le répertoire après avoir vérifié sa frontière.</summary>
            public void Dispose()
            {
                Operations.Dispose();
                string boundary = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "git-branches")) + Path.DirectorySeparatorChar;
                string fullRoot = Path.GetFullPath(Root);
                Guid fixtureId;
                if (!fullRoot.StartsWith(boundary, StringComparison.OrdinalIgnoreCase) ||
                    !string.Equals(Path.GetDirectoryName(fullRoot) + Path.DirectorySeparatorChar, boundary, StringComparison.OrdinalIgnoreCase) ||
                    !Guid.TryParseExact(Path.GetFileName(fullRoot), "N", out fixtureId))
                    throw new InvalidOperationException("Fixture cleanup escaped its own directory");
                // Git object paths can exceed MAX_PATH. Use extended paths for every
                // filesystem call without changing process-wide path handling switches.
                string extendedRoot = fullRoot.StartsWith(@"\\", StringComparison.Ordinal)
                    ? @"\\?\UNC\" + fullRoot.Substring(2) : @"\\?\" + fullRoot;
                if (!Directory.Exists(extendedRoot)) return;
                var pending = new Stack<string>(); var directories = new List<string>(); var files = new List<string>();
                pending.Push(extendedRoot);
                while (pending.Count > 0)
                {
                    string directory = pending.Pop();
                    if ((File.GetAttributes(directory) & FileAttributes.ReparsePoint) != 0)
                        throw new InvalidOperationException("Fixture cleanup refuses filesystem links");
                    directories.Add(directory);
                    foreach (string child in Directory.GetFileSystemEntries(directory, "*", SearchOption.TopDirectoryOnly))
                    {
                        if (!child.StartsWith(extendedRoot + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase))
                            throw new InvalidOperationException("Fixture cleanup escaped its own directory");
                        FileAttributes attributes = File.GetAttributes(child);
                        if ((attributes & FileAttributes.ReparsePoint) != 0)
                            throw new InvalidOperationException("Fixture cleanup refuses filesystem links");
                        if ((attributes & FileAttributes.Directory) != 0) pending.Push(child);
                        else files.Add(child);
                    }
                }
                // Validate the entire owned tree before the first deletion. Failures
                // propagate and remain visible to the test runner.
                foreach (string file in files) { File.SetAttributes(file, FileAttributes.Normal); File.Delete(file); }
                for (int i = directories.Count - 1; i >= 0; i--)
                {
                    File.SetAttributes(directories[i], FileAttributes.Normal);
                    Directory.Delete(directories[i], false);
                }
            }
        }
    }
}
