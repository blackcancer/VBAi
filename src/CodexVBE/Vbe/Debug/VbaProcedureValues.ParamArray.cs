using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace CodexVBE
{
    internal static partial class VbaProcedureValues
    {
        /// <summary>Lit le dernier paramètre VBA ParamArray, implicitement ou explicitement Variant.</summary>
        private static Parameter ReadParamArrayParameter(string[] tokens)
        {
            bool implicitVariant = tokens.Length == 4;
            bool explicitVariant = tokens.Length == 6 && Same(tokens[4], "As") && Same(tokens[5], "Variant");
            if ((!implicitVariant && !explicitVariant) || !Regex.IsMatch(tokens[1], @"^\p{L}[\p{L}\p{N}_]*$") ||
                tokens[2] != "(" || tokens[3] != ")")
                throw new InvalidOperationException("ParamArray requires one unbounded Variant array, without ByVal, ByRef, Optional or a default.");
            return new Parameter { Name = tokens[1], Type = "Variant", ParamArray = true };
        }

        /// <summary>Lie les préfixes ByVal puis conserve chaque valeur du ParamArray comme argument Excel.Run distinct.</summary>
        private static object[] BindParamArrayValues(IReadOnlyList<Parameter> parameters, object[] values, string[] names)
        {
            if (names != null && names.Length > 0)
                throw new ArgumentException("All arguments of a VBA procedure declaring ParamArray must be positional; omit ArgumentNames.");
            int fixedCount = parameters.Count - 1;
            if (values.Length < fixedCount)
                throw new ArgumentException("A required fixed parameter preceding ParamArray is absent: " + parameters[values.Length].Name);
            var bound = new object[values.Length];
            for (int index = 0; index < fixedCount; index++) bound[index] = Coerce(values[index], parameters[index].Type);
            for (int index = fixedCount; index < values.Length; index++) bound[index] = values[index];
            return bound;
        }
    }
}
