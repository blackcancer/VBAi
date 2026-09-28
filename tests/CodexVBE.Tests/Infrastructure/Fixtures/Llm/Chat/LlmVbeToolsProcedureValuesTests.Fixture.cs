namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using System.Linq;

    public sealed partial class LlmVbeToolsBoundaryTests
    {
        /// <summary>Matrice complète des valeurs refusées avant tout appel à la frontière d'exécution.</summary>
        private static IEnumerable<object> RejectedProcedureValues()
        {
            yield return null;
            yield return "not an array";
            yield return new object[31];
            yield return new object[] { new { Member = 1 } };
            yield return new object[] { new object[0] };
            yield return new object[] { new object[] { 1, new object[] { 2 } } };
            yield return new object[] { new object[] { new object[] { 1 }, new object[] { 2, 3 } } };
            yield return new object[] { new object[] { new object[] { new object[] { 1 } } } };
            yield return new object[] { Enumerable.Repeat((object)1, 1025).ToArray() };
            yield return Enumerable.Range(0, 5).Select(_ => (object)Enumerable.Repeat((object)1, 1024).ToArray()).ToArray();
            yield return new object[] { new string('x', 16385) };
            yield return Enumerable.Repeat((object)new string('x', 16384), 5).ToArray();
        }

        /// <summary>Construit une demande JSON munie de toutes les préconditions d'exécution déclarées.</summary>
        private static Dictionary<string, object> ProcedureValuesArguments()
        {
            var values = Arguments("run_procedure_values");
            values["Module"] = "Module1";
            values["Procedure"] = "Values";
            values["ExpectedHostPath"] = @"C:\Temp\fixture.xlsm";
            values["ExpectedSha256"] = "source-revision";
            values["Arguments"] = new object[] { new object[] { "quote\"", null, true, 7 }, new object[] { new object[] { 1, 2 }, new object[] { 3, 4 } } };
            values["ArgumentNames"] = new[] { "vector", "matrix" };
            return values;
        }
    }
}
