namespace CodexVBE
{
    internal sealed partial class LlmVbeTools
    {
        /// <summary>Décrit les arguments JSON scalaires ou tableaux rectangulaires de rang un/deux.</summary>
        private static object ProcedureValuesArgumentSchema()
        {
            var scalars = new object[] { new { type = "string" }, new { type = "number" }, new { type = "boolean" }, new { type = "null" } };
            var scalar = new { anyOf = scalars };
            var vector = new { type = "array", minItems = 1, maxItems = 1024, items = scalar };
            var matrix = new { type = "array", minItems = 1, maxItems = 1024, items = vector };
            return new { type = "array", maxItems = 30, items = new { anyOf = new object[] { scalar, vector, matrix } } };
        }
    }
}
