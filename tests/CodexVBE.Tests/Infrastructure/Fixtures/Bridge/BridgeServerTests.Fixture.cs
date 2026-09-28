namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.IO.Pipes;
    using System.Text;
    using System.Threading;
    using System.Threading.Tasks;
    using System.Web.Script.Serialization;
    using System.Windows.Forms;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>Helpers partagés par les tests du serveur de pont IPC.</summary>
    public sealed partial class BridgeServerTests
    {
        /// <summary>Envoie une requête au canal nommé en pompant la boucle WinForms du thread VBE.</summary>
        /// <param name="processId">Identifiant utilisé pour nommer le canal serveur.</param>
        /// <param name="request">Ligne JSON transmise au serveur.</param>
        /// <returns>Réponse JSON désérialisée.</returns>
        /// <exception cref="AssertFailedException">La réponse n’est pas reçue dans le délai de dix secondes.</exception>
        private static IDictionary<string, object> SendWithMessagePump(int processId, string request)
        {
            var pending = Task.Run(() =>
            {
                using (var pipe = new NamedPipeClientStream(".", "CodexVBE." + processId, PipeDirection.InOut))
                {
                    pipe.Connect(5000);
                    using (var writer = new StreamWriter(pipe, new UTF8Encoding(false), 4096, true)
                    {
                        AutoFlush = true
                    }

                    )
                    using (var reader = new StreamReader(pipe, new UTF8Encoding(false), false, 4096, true))
                    {
                        writer.WriteLine(request);
                        return (IDictionary<string, object>)new JavaScriptSerializer().DeserializeObject(reader.ReadLine());
                    }
                }
            });
            var deadline = DateTime.UtcNow.AddSeconds(10);
            while (!pending.IsCompleted && DateTime.UtcNow < deadline)
            {
                Application.DoEvents();
                Thread.Sleep(5);
            }

            Assert.IsTrue(pending.IsCompleted, "The VBE bridge did not complete a pipe request.");
            return pending.GetAwaiter().GetResult();
        }

        /// <summary>Hôte VBE simulé avec sa collection de projets.</summary>
        public sealed class FakeVbe
        {
            /// <summary>Projets visibles par les tests du pont.</summary>
            /// <value>Liste des projets factices du VBE.</value>
            public List<FakeProject> VBProjects { get; } = new List<FakeProject>();
        }

        /// <summary>Projet VBE factice avec identité de document et mode courant.</summary>
        public sealed class FakeProject
        {
            /// <summary>Nom affiché dans le VBE.</summary>
            /// <value>Nom du projet.</value>
            public string Name { get; set; }
            /// <summary>Chemin du document associé.</summary>
            /// <value>Chemin du classeur ou chaîne vide pour un document non enregistré.</value>
            public string FileName { get; set; }
            /// <summary>Mode du projet, initialisé au mode création.</summary>
            /// <value>Valeur du mode VBE simulé.</value>
            public int Mode { get; set; } = 2;
        }
    }
}
