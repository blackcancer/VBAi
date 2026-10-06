using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace VBAi
{

    /// <summary>Source candidates only: never claims a complete runtime Locals inventory.</summary>
    internal static class VbaLocalScalarCandidates
    {

        /// <summary>Owns the candidate state and operations.</summary>
        internal sealed class Candidate
        {

            /// <summary>Maintains the name and expression and kind and type name and reason state for candidate.</summary>
            public string Name, Expression, Kind, TypeName, Reason;

            /// <summary>Maintains the line and column state for candidate.</summary>
            public int Line, Column;

            /// <summary>Maintains the eligible state for candidate.</summary>
            public bool Eligible;
        }

        /// <summary>Maintains the scalar types state for vba local scalar candidates.</summary>
        private static readonly HashSet<string> ScalarTypes = new HashSet<string>(StringComparer.OrdinalIgnoreCase) {
            "Boolean", "Byte", "Integer", "Long", "LongLong", "LongPtr", "Single", "Double", "Currency", "Date", "String"
        };

        /// <summary>Reads  for vba local scalar candidates.</summary>
        /// <param name="source">Text that supplies the source value. Use the format required by the calling operation.</param>
        /// <param name="procedure">Text that supplies the procedure value. Use the format required by the calling operation.</param>
        /// <param name="first">int that supplies the first for this operation.</param>
        /// <param name="last">int that supplies the last for this operation.</param>
        /// <returns>candidate[] produced by the operation for read on vba local scalar candidates.</returns>
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
