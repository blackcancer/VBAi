using System;
using System.Collections.Generic;
using CodexVBE;

namespace CodexVBE.Tests.Unit
{
    public sealed partial class VbeDebugTests
    {
        /// <summary>Matrice établie avant tests: scalar/vector/matrix/null, copies, binding, bornes, rejets et unique invocation.</summary>
        private static readonly string[] ValueSignatureFailures = { "ByRef", "implicit ByRef", "typed array", "object", "invalid ParamArray", "conditional", "private", "property", "missing required", "unknown named", "duplicate named", "return object" };
        /// <summary>Gardes indépendantes avant mise en file ou livraison native.</summary>
        private static readonly string[] ValueQueueFailures = { "stale source", "wrong path", "missing path", "runtime mode", "class target", "host identity", "changed source", "changed path", "changed mode" };

        /// <summary>Transport injecté sans hôte COM, conservant le nombre exact d'invocations et les arguments transmis.</summary>
        private sealed class ValuesHost : VbeDebug.IProcedureValuesHost
        {
            internal int Resolves, Invocations;
            internal bool RejectIdentity;
            internal object Return = new object[] { 7, "ok", true, DBNull.Value };
            internal string Module, Procedure;
            internal object[] Received;
            internal Action OnInvoke;
            internal Action OnResolve;
            public object ResolveTarget(object project, string expectedHostPath)
            { Resolves++; if (RejectIdentity) throw new InvalidOperationException("Owned PID/COM identity unavailable"); OnResolve?.Invoke(); return this; }
            public object Invoke(object target, string module, string procedure, object[] arguments)
            { Invocations++; Module = module; Procedure = procedure; Received = arguments; OnInvoke?.Invoke(); return Return; }
        }

        /// <summary>Prépare une signature Variant pour les tableaux avec chemin de classeur inspecté.</summary>
        private static Request ValuesRequest(Fixture fixture, string signature = "Public Function TryMe(ByVal values As Variant) As Variant")
        {
            fixture.Project.VBComponents[0].Type = 1;
            fixture.Module.Code = signature + "\r\nTryMe = values\r\nEnd Function";
            var request = Location(fixture); request.Procedure = "TryMe"; request.ExpectedSha256 = Sha(fixture.Module.Code);
            request.ExpectedHostPath = fixture.Project.FileName; request.Arguments = new object[] { new object[] { 1, "quoted\" value", null } };
            return request;
        }
    }
}
