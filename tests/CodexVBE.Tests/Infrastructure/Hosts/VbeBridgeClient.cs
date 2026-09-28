using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Pipes;
using System.Text;
using System.Threading;
using System.Web.Script.Serialization;

namespace CodexVBE.Tests.Integration
{
    /// <summary>Envoie des commandes JSON au serveur de test nommé du complément VBE.</summary>
    internal static class VbeBridgeClient
    {
        /// <summary>Convertit un objet désérialisé en dictionnaire de propriétés JSON.</summary>
        /// <param name="value">Valeur produite par le désérialiseur.</param>
        /// <returns>Le dictionnaire de propriétés de l’objet JSON.</returns>
        internal static IDictionary<string, object> Object(object value)
        {
            return (IDictionary<string, object>)value;
        }

        /// <summary>Envoie une commande textuelle au canal nommé du processus.</summary>
        /// <param name="processId">Identifiant du processus qui héberge le canal.</param>
        /// <param name="command">Nom de la commande à envoyer.</param>
        /// <returns>La réponse convertie en dictionnaire, ou <see langword="null"/> après des tentatives échouées.</returns>
        internal static IDictionary<string, object> Read(int processId, string command)
        {
            return Read(processId, new { Command = command });
        }

        /// <summary>Envoie une requête au canal nommé et lit sa réponse JSON en réessayant après les échecs transitoires.</summary>
        /// <param name="processId">Identifiant du processus qui héberge le canal.</param>
        /// <param name="request">Objet sérialisable envoyé sous forme d’une ligne JSON.</param>
        /// <returns>La réponse convertie en dictionnaire, ou <see langword="null"/> après épuisement des tentatives.</returns>
        internal static IDictionary<string, object> Read(int processId, object request)
        {
            var json = new JavaScriptSerializer();
            for (int attempt = 0; attempt < 20; attempt++)
            {
                try
                {
                    using (var pipe = new NamedPipeClientStream(".", "CodexVBE." + processId, PipeDirection.InOut))
                    {
                        pipe.Connect(250);
                        using (var writer = new StreamWriter(pipe, new UTF8Encoding(false), 4096, true) { AutoFlush = true })
                        using (var reader = new StreamReader(pipe, new UTF8Encoding(false), false, 4096, true))
                        {
                            writer.WriteLine(json.Serialize(request));
                            return Object(json.DeserializeObject(reader.ReadLine()));
                        }
                    }
                }
                catch (TimeoutException) { Thread.Sleep(250); }
                catch (IOException) { Thread.Sleep(250); }
            }
            return null;
        }
    }
}
