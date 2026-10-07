using System;

namespace VBAi.Tests.Unit
{
    /// <summary>Matrice écrite avant les tests: résolution, exclusion, refus et concurrence du projet entier.</summary>
    internal static class ProcedureRenameMatrix
    {
        /// <summary>Déclaration cible, récursion et variable de retour.</summary>
        internal const string Target = "Option Explicit\r\nPublic Function Calc(ByVal value As Long) As Long\r\nCalc = Calc(value - 1)\r\nDebug.Print \"Calc\" ' Calc\r\nEnd Function";
        /// <summary>Appels directs, qualifiés, chaînes, membres COM, arguments nommés et blocs With.</summary>
        internal const string Caller = "Option Explicit\r\nPublic Sub Caller()\r\nDebug.Print Calc(1), MathModule.Calc(2), P.MathModule.Calc(3), obj.Calc(4)\r\nOther Calc:=1\r\nWith obj\r\n.Calc(5)\r\nEnd With\r\nGoTo Calc\r\nCalc: Debug.Print \"Calc\"\r\nEnd Sub";
        /// <summary>Refus à source complète avant toute mutation.</summary>
        internal static readonly string[] Refusals = { "class target", "missing target", "wrong position", "stale sha", "property kind", "duplicate module",
            "procedure collision", "new procedure collision", "local shadow", "qualifier shadow", "new implicit binding", "no explicit", "conditional", "dynamic call",
            "callback", "interface", "bracket expression", "unterminated", "external qualification", "private caller", "new callback", "reserved name" };
        /// <summary>Prépare l'identité de déclaration et les gardes de version du module.</summary>
        internal static Request Request(string source = Target) => new Request
        {
            Project = "P",
            Module = "MathModule",
            Query = "Calc",
            Procedure = "Calc",
            NewName = "Compute",
            StartLine = 2,
            StartColumn = source.Split('\n')[1].IndexOf("Calc", StringComparison.Ordinal) + 1,
            ExpectedSha256 = VbaProcedureRename.Digest(source),
            ExpectedMode = 2
        };
        /// <summary>Construit l'instantané minimal avec un module standard et un appelant classe.</summary>
        internal static VbaProcedureRename.ModuleSnapshot[] Project(string target = Target, string caller = Caller, int targetType = 1) =>
            new[] { new VbaProcedureRename.ModuleSnapshot("MathModule", targetType, target), new VbaProcedureRename.ModuleSnapshot("Caller", 2, caller) };
    }
}
