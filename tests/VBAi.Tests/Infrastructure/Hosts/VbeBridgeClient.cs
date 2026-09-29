using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Pipes;
using System.Text;
using System.Threading;
using System.Web.Script.Serialization;

namespace VBAi.Tests.Integration
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

        /// <summary>Réessaie uniquement la connexion ; toute erreur après début d'émission interdit une réémission.</summary>
        /// <param name="processId">Identifiant du processus propriétaire du canal.</param>
        /// <param name="request">Objet sérialisable envoyé une seule fois.</param>
        /// <returns>Réponse JSON, ou null si aucune connexion n'a pu être établie.</returns>
        internal static IDictionary<string, object> Read(int processId, object request) =>
            Read("VBAi." + processId, request, 120000);

        /// <summary>Transport borné injectable pour des pipes de tests possédés, sans retry après WriteLine.</summary>
        /// <param name="pipeName">Nom exact du seul canal à connecter.</param>
        /// <param name="request">Requête sérialisée avant toute connexion.</param>
        /// <param name="responseTimeoutMilliseconds">Temps maximal de lecture après émission ; défaut 120 secondes.</param>
        /// <param name="connectTimeoutMilliseconds">Temps maximal par tentative de connexion.</param>
        /// <param name="connectionAttempts">Nombre borné de tentatives avant émission seulement.</param>
        /// <param name="retryDelayMilliseconds">Pause bornée entre échecs de connexion.</param>
        /// <param name="onConnectionRetry">Observation de connexion échouée pour les fixtures déterministes ; appelée avant toute émission.</param>
        /// <returns>Réponse complète, ou null après épuisement des seules connexions.</returns>
        internal static IDictionary<string, object> Read(string pipeName, object request, int responseTimeoutMilliseconds = 120000,
            int connectTimeoutMilliseconds = 250, int connectionAttempts = 20, int retryDelayMilliseconds = 250, Action<int> onConnectionRetry = null)
        {
            if (string.IsNullOrWhiteSpace(pipeName) || responseTimeoutMilliseconds < 1 || responseTimeoutMilliseconds > 600000 ||
                connectTimeoutMilliseconds < 1 || connectTimeoutMilliseconds > 10000 || connectionAttempts < 1 || connectionAttempts > 100 ||
                retryDelayMilliseconds < 0 || retryDelayMilliseconds > 10000)
                throw new ArgumentException("A named pipe and bounded positive response/connection timeouts are required.");
            var json = new JavaScriptSerializer { MaxJsonLength = 10 * 1024 * 1024 };
            string payload = json.Serialize(request);
            for (int attempt = 0; attempt < connectionAttempts; attempt++)
            {
                var pipe = new NamedPipeClientStream(".", pipeName, PipeDirection.InOut, PipeOptions.Asynchronous);
                bool connected = false;
                try
                {
                    try { pipe.Connect(connectTimeoutMilliseconds); connected = true; }
                    catch (TimeoutException) { }
                    catch (IOException) { }
                    if (!connected)
                    {
                        if (attempt + 1 < connectionAttempts)
                        {
                            onConnectionRetry?.Invoke(attempt + 1);
                            if (retryDelayMilliseconds > 0) Thread.Sleep(retryDelayMilliseconds);
                        }
                        continue;
                    }
                    // From this call onward the request may have reached the host. No failure below is retried.
                    using (var writer = new StreamWriter(pipe, new UTF8Encoding(false), 4096, true) { AutoFlush = true })
                        writer.WriteLine(payload);
                    // Dispose the flushed writer before reading so a broken-pipe flush cannot mask a response error.
                    using (var reader = new StreamReader(pipe, new UTF8Encoding(false), false, 4096, true))
                    {
                        var reading = reader.ReadLineAsync();
                        using (var delayCancellation = new CancellationTokenSource())
                        {
                            var deadline = System.Threading.Tasks.Task.Delay(responseTimeoutMilliseconds, delayCancellation.Token);
                            var completed = System.Threading.Tasks.Task.WhenAny(reading, deadline).GetAwaiter().GetResult();
                            if (completed != reading)
                            {
                                // The finally closes the owned client and cancels pending I/O. Observe any fault without waiting again.
                                reading.ContinueWith(task => task.Exception.Handle(error => true),
                                    System.Threading.Tasks.TaskContinuationOptions.OnlyOnFaulted);
                                throw new TimeoutException("The VBE response deadline expired after emission; delivery is uncertain and the request was not retried.");
                            }
                            delayCancellation.Cancel();
                        }
                        string line = reading.GetAwaiter().GetResult();
                        if (line == null) throw new IOException("The VBE pipe closed after emission; delivery is uncertain and the request was not retried.");
                        object parsed = json.DeserializeObject(line);
                        if (!(parsed is IDictionary<string, object> response))
                            throw new InvalidDataException("The VBE response is not a JSON object; the emitted request was not retried.");
                        return response;
                    }
                }
                finally { pipe.Dispose(); }
            }
            return null;
        }
    }
}