namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using System.Threading;
    using System.Web.Script.Serialization;
    using System.Threading.Tasks;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>Fournit la désérialisation et les assertions communes aux validations asynchrones des outils VBE.</summary>
    public sealed partial class LlmVbeAsyncValidationTests
    {
        /// <summary>Sérialiseur JSON utilisé pour lire les réponses d’outil.</summary>
        private static readonly JavaScriptSerializer Json = new JavaScriptSerializer();
        /// <summary>Vérifie qu’un appel d’outil échoue avec le fragment de message attendu.</summary>
        /// <param name="tools">Outils VBE à invoquer.</param>
        /// <param name="name">Nom de l’outil.</param>
        /// <param name="arguments">Arguments JSON de l’outil.</param>
        /// <param name="fragment">Fragment attendu dans le message d’erreur.</param>
        private static async Task Failure(LlmVbeTools tools, string name, string arguments, string fragment)
        {
            string serialized = await tools.InvokeAsync(name, arguments);
            var response = Json.Deserialize<Response>(serialized);
            Assert.IsFalse(response.Ok, name + " accepted invalid arguments");
            StringAssert.Contains(response.Error, fragment);
        }
    }
}
