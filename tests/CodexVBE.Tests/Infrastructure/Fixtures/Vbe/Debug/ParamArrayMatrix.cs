using System.Collections.Generic;

namespace CodexVBE.Tests.Infrastructure
{
    /// <summary>Signatures ParamArray invalides préparées avant l'exécution de la matrice.</summary>
    internal static class ParamArrayMatrix
    {
        internal static IEnumerable<string> RejectedSignatures()
        {
            yield return "ParamArray values As Variant";
            yield return "ParamArray values(1) As Variant";
            yield return "ParamArray values() As Long";
            yield return "ParamArray values() As Object";
            yield return "ParamArray values() As Variant = 0";
            yield return "ParamArray values(), ByVal after As Long";
            yield return "ParamArray first(), ParamArray second()";
            yield return "Optional ByVal first As Variant, ParamArray values()";
            yield return "Optional ParamArray values() As Variant";
            yield return "ByVal ParamArray values() As Variant";
            yield return "ByRef ParamArray values() As Variant";
            yield return "ByVal first As Long, ParamArray FIRST() As Variant";
            yield return "ParamArray values!() As Variant";
        }
    }
}
