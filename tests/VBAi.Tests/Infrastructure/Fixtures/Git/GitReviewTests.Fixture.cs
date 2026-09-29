namespace VBAi.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Net;
    using System.Net.Http;
    using System.Threading;
    using System.Threading.Tasks;
    using System.Web.Script.Serialization;
    using VBAi;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>Contient les doubles HTTP utilisés par les tests de revue GitHub.</summary>
    public sealed partial class GitReviewTests
    {
        /// <summary>Transmet les requêtes à un délégué configurable sans réseau réel.</summary>
        private sealed class Handler : HttpMessageHandler
        {
            /// <summary>Fonction qui produit la réponse associée à chaque requête.</summary>
            internal Func<HttpRequestMessage, Task<HttpResponseMessage>> Send;
            /// <summary>Vérifie l’annulation puis transmet la requête au délégué configuré.</summary>
            /// <param name="request">Requête HTTP reçue par le gestionnaire.</param>
            /// <param name="token">Jeton d’annulation de l’envoi.</param>
            /// <returns>La réponse asynchrone produite par le délégué.</returns>
            protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken token)
            {
                token.ThrowIfCancellationRequested();
                return Send(request);
            }
        }
    }
}
