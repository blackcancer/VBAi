namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Web.Script.Serialization;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>Fournit les utilitaires JSON communs aux tests de contrat des outils VBE.</summary>
    public sealed partial class LlmVbeToolContractTests
    {
        /// <summary>Sérialiseur configuré pour accepter les réponses volumineuses des outils.</summary>
        private static readonly JavaScriptSerializer Json = new JavaScriptSerializer
        {
            MaxJsonLength = 10 * 1024 * 1024
        };
        /// <summary>Convertit un objet JSON en dictionnaire clé-valeur.</summary>
        /// <param name="value">Valeur désérialisée.</param>
        /// <returns>Dictionnaire représentant l’objet JSON.</returns>
        private static IDictionary<string, object> Dict(object value)
        {
            return (IDictionary<string, object>)value;
        }

        /// <summary>Vérifie qu’une réponse JSON est un échec contenant le texte indiqué.</summary>
        /// <param name="serialized">Réponse JSON sérialisée.</param>
        /// <param name="fragment">Fragment attendu dans le message d’erreur.</param>
        private static void IsFailure(string serialized, string fragment)
        {
            var response = Json.Deserialize<Response>(serialized);
            Assert.IsFalse(response.Ok);
            StringAssert.Contains(response.Error, fragment);
        }
    }
}
