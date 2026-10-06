using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace VBAi
{

    /// <summary>Source candidates only: never claims a complete runtime Locals inventory.</summary>
    internal static class VbaLocalScalarCandidates
    {

        /// <summary>One source-declared local or parameter candidate, with source coordinates and eligibility diagnostics.</summary>
        internal sealed class Candidate
        {

            /// <summary>Name/token expression, declaration kind, declared VBA type, and exclusion reason when ineligible.</summary>
            public string Name, Expression, Kind, TypeName, Reason;

            /// <summary>One-based source line and column of the declaration name token.</summary>
            public int Line, Column;

            /// <summary>True only for an unambiguous, supported, unconditional scalar declaration.</summary>
            public bool Eligible;
        }

        /// <summary>VBA scalar type names supported by the source-based local candidate filter.</summary>
        private static readonly HashSet<string> ScalarTypes = new HashSet<string>(StringComparer.OrdinalIgnoreCase) {
            "Boolean", "Byte", "Integer", "Long", "LongLong", "LongPtr", "Single", "Double", "Currency", "Date", "String"
        };

        /// <summary>Extracts source declarations in one procedure/range and marks conditional, duplicate, array, auto-instanced, and unsupported types as ineligible.</summary>
        /// <param name="source">VBA module source to tokenize and index.</param>
        /// <param name="procedure">Procedure name whose local declarations are selected, compared case-insensitively.</param>
        /// <param name="first">First one-based source line in the inspected range.</param>
        /// <param name="last">Last one-based source line; must be at least <paramref name="first"/>.</param>
        /// <returns>Candidate records only; this source scan is not a complete runtime Locals inventory.</returns>
        internal static Candidate[] Read(string source, string procedure, int first, int last)
        {
            if (source == null || string.IsNullOrWhiteSpace(procedure) || first < 1 || last < first)
                throw new ArgumentException("Source and a valid procedure range are required.");
            var statements = VbaDeclarationIndex.Statements(source).ToArray();
            var locations = statements.SelectMany(s => s.Select((token, index) => new { Token = token, Statement = s, Index = index }))
                .ToDictionary(item => item.Token.Line + ":" + item.Token.Column);
            var declarations = VbaDeclarationIndex.Read(source).Where(d =>
                d.Line >= first && d.Line <= last && string.Equals(d.Scope, procedure, StringComparison.OrdinalIgnoreCase) &&
                (d.Kind == "Variable" || d.Kind == "Parameter" || d.Kind == "Constant")).ToArray();
            var duplicateNames = new HashSet<string>(declarations.GroupBy(d => d.Name, StringComparer.OrdinalIgnoreCase)
                .Where(g => g.Count() > 1).Select(g => g.Key), StringComparer.OrdinalIgnoreCase);
            return declarations.Select(d => {
                locations.TryGetValue(d.Line + ":" + d.Column, out var location);
                var token = location?.Token;
                var statement = location?.Statement;
                int index = location?.Index ?? -1;
                string reason = d.Conditional ? "ConditionalDeclaration" : duplicateNames.Contains(d.Name) ? "AmbiguousDeclaration" :
                    token == null || !Regex.IsMatch(token.Text, @"^\p{L}[\p{L}\p{N}_]*[$%&!#@^]?$") ? "UnsupportedIdentifier" :
                    index + 1 < statement.Count && statement[index + 1].Text == "(" ? "ArrayDeclaration" :
                    statement.Skip(index + 1).TakeWhile(t => t.Text != ",").Any(t => t.Text.Equals("New", StringComparison.OrdinalIgnoreCase)) ? "AutoInstantiatedDeclaration" :
                    !ScalarTypes.Contains(d.TypeName ?? "") ? "NonScalarOrVariantType" : null;
                return new Candidate { Name = d.Name, Expression = token?.Text, Kind = d.Kind, TypeName = d.TypeName,
                    Line = d.Line, Column = d.Column, Eligible = reason == null, Reason = reason };
            }).ToArray();
        }
    }
}
