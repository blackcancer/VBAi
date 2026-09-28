using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

namespace CodexVBE
{
    /// <summary>Liaison bornée de valeurs JSON aux paramètres ByVal et normalisation des retours scalaires/SAFEARRAY.</summary>
    internal static partial class VbaProcedureValues
    {
        /// <summary>Résultat sérialisable conservant les bornes natives des tableaux retournés.</summary>
        internal sealed class Result
        {
                        /// <summary>Scalar, Null, Empty ou Array.</summary>
                        /// <value>Catégorie sérialisable du résultat retourné par la procédure.</value>
            public string Kind { get; }
                        /// <summary>Scalaire JSON ou tableau JSON; aucune référence COM.</summary>
                        /// <value>Valeur normalisée, ou nul pour les résultats Null et Empty.</value>
            public object Value { get; }
                        /// <summary>Bornes basses natives, absentes pour un scalaire.</summary>
                        /// <value>Bornes de chaque dimension du SAFEARRAY, ou tableau vide pour un scalaire.</value>
            public int[] LowerBounds { get; }
                        /// <summary>Longueurs des dimensions natives.</summary>
                        /// <value>Nombre d’éléments de chaque dimension du tableau natif.</value>
            public int[] Lengths { get; }
                        /// <summary>Construit un résultat après validation de tous les éléments.</summary>
                        /// <param name="kind">Catégorie du résultat.</param>
                        /// <param name="value">Valeur scalaire ou tableau convertie pour JSON.</param>
                        /// <param name="lowerBounds">Bornes basses natives de chaque dimension.</param>
                        /// <param name="lengths">Longueurs natives de chaque dimension.</param>
            internal Result(string kind, object value, int[] lowerBounds, int[] lengths)
            { Kind = kind; Value = value; LowerBounds = lowerBounds; Lengths = lengths; }
        }

        /// <summary>Paramètre de signature explicitement ByVal; aucun tableau typé ou objet n'est accepté.</summary>
        private sealed class Parameter
        {
            /// <summary>Nom et type déclarés du paramètre VBA.</summary>
            internal string Name, Type;
            /// <summary>Indique si le paramètre est facultatif ou un ParamArray.</summary>
            internal bool Optional, ParamArray;
        }

                /// <summary>Copie profonde des arguments pour empêcher toute mutation de la requête après mise en file.</summary>
                /// <param name="arguments">Valeurs JSON scalaires ou tableaux rectangulaires à copier.</param>
                /// <returns>Arguments normalisés sous forme de variantes ou tableaux de variantes.</returns>
        internal static object[] Capture(object[] arguments)
        {
            if (arguments == null || arguments.Length > 30) throw new ArgumentException("At most 30 JSON Arguments are required.");
            int cells = 0, text = 0;
            return arguments.Select(value => CaptureValue(value, ref cells, ref text)).ToArray();
        }

                /// <summary>Lie les noms à l'ordre de déclaration; les omissions Optional deviennent Type.Missing pour Excel.Run.</summary>
                /// <param name="source">Texte du module dont la signature a été relue.</param>
                /// <param name="bodyLine">Ligne exacte du corps de la procédure dans ce texte.</param>
                /// <param name="procedure">Nom de la procédure cible.</param>
                /// <param name="values">Arguments capturés fournis par l’appelant.</param>
                /// <param name="names">Noms d’arguments facultatifs, ou nul pour un appel positionnel.</param>
                /// <returns>Arguments convertis et ordonnés pour l’invocation Excel.Run.</returns>
        internal static object[] Bind(string source, int bodyLine, string procedure, object[] values, string[] names)
        {
            var statements = VbaDeclarationIndex.Statements(source).ToArray();
            if (statements.Any(x => x.Count > 0 && (x[0].Text == "#" || Regex.IsMatch(x[0].Text, @"^Def(?:Bool|Byte|Int|Lng|LngLng|LngPtr|Sng|Dbl|Cur|Date|Str|Obj|Var)$", RegexOptions.IgnoreCase))))
                throw new InvalidOperationException("Conditional and implicit-type signatures are unsupported.");
            var header = statements.SingleOrDefault(x => x.Count > 0 && x[0].Line == bodyLine);
            if (header == null) throw new InvalidOperationException("The exact live signature is absent.");
            int signatures = 0;
            foreach (var statement in statements)
            {
                int at = 0;
                while (at < statement.Count && new[] { "public", "private", "friend", "static" }.Contains(statement[at].Text.ToLowerInvariant())) at++;
                if (at + 1 < statement.Count && (Same(statement[at].Text, "Sub") || Same(statement[at].Text, "Function")) && Same(statement[at + 1].Text, procedure)) signatures++;
            }
            if (signatures != 1) throw new InvalidOperationException("The procedure signature is repeated or unresolved.");
            int first = 0;
            while (first < header.Count && (Same(header[first].Text, "Public") || Same(header[first].Text, "Static"))) first++;
            if (first + 2 >= header.Count || (!Same(header[first].Text, "Sub") && !Same(header[first].Text, "Function")) ||
                !Same(header[first + 1].Text, procedure) || header[first + 2].Text != "(")
                throw new InvalidOperationException("Only an exact Public/default-public standard-module Sub/Function is callable.");
            int opening = first + 2, closing = -1, depth = 0;
            for (int i = opening; i < header.Count; i++)
            { if (header[i].Text == "(") depth++; else if (header[i].Text == ")" && --depth == 0) { closing = i; break; } }
            if (closing < 0) throw new InvalidOperationException("The signature parentheses are incomplete.");
            if (Same(header[first].Text, "Function") && closing + 1 < header.Count)
            {
                int suffixCount = header.Count - closing - 1;
                bool arraySuffix = suffixCount == 4 && header[closing + 3].Text == "(" && header[closing + 4].Text == ")";
                if (!Same(header[closing + 1].Text, "As") || closing + 2 >= header.Count || !ScalarType(header[closing + 2].Text) ||
                    (suffixCount != 2 && !arraySuffix))
                    throw new InvalidOperationException("Object, class, UDT, Date and other non-JSON function return types are unsupported.");
            }
            var parameters = new List<Parameter>(); int start = opening + 1; depth = 0;
            for (int i = start; i <= closing; i++)
            {
                if (i < closing && header[i].Text == "(") depth++;
                if (i < closing && header[i].Text == ")") depth--;
                if (i < closing && (header[i].Text != "," || depth != 0)) continue;
                var tokens = header.Skip(start).Take(i - start).Select(x => x.Text).ToArray(); start = i + 1;
                if (tokens.Length == 0) { if (i != opening + 1) throw new InvalidOperationException("An empty parameter is unreadable."); continue; }
                if (Same(tokens[0], "ParamArray"))
                {
                    if (i != closing || parameters.Any(x => x.Optional))
                        throw new InvalidOperationException("ParamArray must be last and cannot share a signature with Optional parameters.");
                    parameters.Add(ReadParamArrayParameter(tokens));
                    continue;
                }
                int at = 0; bool optional = Same(tokens[at], "Optional"); if (optional) at++;
                if (at >= tokens.Length || !Same(tokens[at++], "ByVal") || at >= tokens.Length)
                    throw new InvalidOperationException("Every fixed signature parameter must be explicitly ByVal; ByRef/default-ByRef are unsupported.");
                string name = tokens[at++];
                if (!Regex.IsMatch(name, @"^\p{L}[\p{L}\p{N}_]*$") || at + 1 >= tokens.Length || !Same(tokens[at++], "As") || !ScalarType(tokens[at]))
                    throw new InvalidOperationException("Typed array parameters, objects/classes/UDTs and implicit parameter types are unsupported; JSON arrays require ByVal As Variant.");
                string type = tokens[at++];
                if (at < tokens.Length && tokens[at] != "=") throw new InvalidOperationException("The parameter suffix is unsupported.");
                if (at < tokens.Length && !optional) throw new InvalidOperationException("Only Optional parameters can declare defaults.");
                parameters.Add(new Parameter { Name = name, Type = type, Optional = optional });
            }
            int fixedParameterCount = parameters.Count - (parameters.Count > 0 && parameters[parameters.Count - 1].ParamArray ? 1 : 0);
            if (fixedParameterCount > 30 || parameters.GroupBy(x => x.Name, StringComparer.OrdinalIgnoreCase).Any(x => x.Count() != 1))
                throw new InvalidOperationException("The signature has too many or duplicate parameters.");
            if (values == null || values.Length > 30 || (names != null && (names.Length != values.Length || names.Distinct(StringComparer.OrdinalIgnoreCase).Count() != names.Length)))
                throw new ArgumentException("ArgumentNames must match Arguments and be distinct.");
            if (parameters.Count > 0 && parameters[parameters.Count - 1].ParamArray)
                return BindParamArrayValues(parameters, values, names);
            var bound = Enumerable.Repeat<object>(Type.Missing, parameters.Count).ToArray();
            for (int i = 0; i < values.Length; i++)
            {
                int index = names == null ? i : parameters.FindIndex(x => Same(x.Name, names[i]));
                if (index < 0 || index >= parameters.Count) throw new ArgumentException("An argument does not belong to the live signature.");
                bound[index] = Coerce(values[i], parameters[index].Type);
            }
            for (int i = 0; i < parameters.Count; i++)
                if (ReferenceEquals(bound[i], Type.Missing) && !parameters[i].Optional) throw new ArgumentException("A required argument is absent: " + parameters[i].Name);
            return bound;
        }

                /// <summary>Convertit une seule fois le retour natif sans invoquer de getter d'objet COM.</summary>
                /// <param name="raw">Valeur retournée par la procédure appelée.</param>
                /// <returns>Résultat JSON normalisé avec les bornes natives des tableaux, le cas échéant.</returns>
        internal static Result NormalizeReturn(object raw)
        {
            if (raw == null) return new Result("Empty", null, new int[0], new int[0]);
            if (raw == DBNull.Value) return new Result("Null", null, new int[0], new int[0]);
            int cells = 0, text = 0;
            if (!(raw is Array array)) return new Result("Scalar", JsonScalar(raw, ref cells, ref text), new int[0], new int[0]);
            if (array.Rank < 1 || array.Rank > 2 || array.LongLength > 4096) throw new InvalidOperationException("Returned arrays support rank 1/2 and at most 4096 scalar cells.");
            int[] lower = Enumerable.Range(0, array.Rank).Select(array.GetLowerBound).ToArray();
            int[] lengths = Enumerable.Range(0, array.Rank).Select(array.GetLength).ToArray();
            var rows = new object[array.GetLength(0)];
            for (int row = 0; row < rows.Length; row++)
            {
                if (array.Rank == 1) rows[row] = JsonScalar(array.GetValue(row + lower[0]), ref cells, ref text);
                else
                {
                    var columns = new object[lengths[1]];
                    for (int column = 0; column < columns.Length; column++) columns[column] = JsonScalar(array.GetValue(row + lower[0], column + lower[1]), ref cells, ref text);
                    rows[row] = columns;
                }
            }
            return new Result("Array", rows, lower, lengths);
        }

                /// <summary>Capture un scalaire ou un tableau JSON rectangulaire rank 1/2; les bornes entrantes commencent à zéro.</summary>
                /// <param name="value">Valeur scalaire ou tableau d’arguments à capturer.</param>
                /// <param name="cells">Compteur partagé des cellules scalaires observées.</param>
                /// <param name="text">Compteur partagé du nombre de caractères copiés.</param>
                /// <returns>Variante scalaire, vecteur ou matrice bidimensionnelle copiée.</returns>
        private static object CaptureValue(object value, ref int cells, ref int text)
        {
            if (!(value is Array array)) return NativeScalar(value, ref cells, ref text);
            if (array.Rank != 1 || array.GetLowerBound(0) != 0 || array.Length == 0 || array.Length > 1024) throw new ArgumentException("JSON arrays must be nonempty zero-based vectors or rectangular nested vectors with at most 1024 cells.");
            var values = array.Cast<object>().ToArray();
            if (values.Any(x => x is Array))
            {
                if (values.Any(x => !(x is Array))) throw new ArgumentException("Mixed scalar/array and ragged arrays are unsupported.");
                var rows = values.Cast<Array>().ToArray(); int columns = rows[0].Length;
                if (columns == 0 || rows.Any(x => x.Rank != 1 || x.GetLowerBound(0) != 0 || x.Length != columns) || (long)rows.Length * columns > 1024)
                    throw new ArgumentException("JSON matrices must be rectangular, zero-based and contain at most 1024 scalar cells.");
                var result = new object[rows.Length, columns];
                for (int row = 0; row < rows.Length; row++)
                    for (int column = 0; column < columns; column++) result[row, column] = NativeScalar(rows[row].GetValue(column), ref cells, ref text);
                return result;
            }
            var vector = new object[values.Length];
            for (int i = 0; i < values.Length; i++) vector[i] = NativeScalar(values[i], ref cells, ref text);
            return vector;
        }

                /// <summary>Normalise les scalaires JSON en sous-types VARIANT portables sans conversion silencieuse de texte/nombre.</summary>
                /// <param name="value">Scalaire à convertir.</param>
                /// <param name="cells">Compteur partagé des cellules.</param>
                /// <param name="text">Compteur partagé des caractères.</param>
                /// <returns>Sous-type scalaire compatible avec VARIANT.</returns>
        private static object NativeScalar(object value, ref int cells, ref int text)
        {
            if (++cells > 4096) throw new ArgumentException("The call exceeds 4096 scalar input cells.");
            if (value == null || value == DBNull.Value) return DBNull.Value;
            if (value is bool) return value;
            if (value is string stringValue)
            { text += stringValue.Length; if (stringValue.Length > 16384 || text > 65536) throw new ArgumentException("Procedure strings exceed the bounded payload."); return value; }
            if (value is double || value is float)
            { double number = Convert.ToDouble(value, CultureInfo.InvariantCulture); if (double.IsNaN(number) || double.IsInfinity(number)) throw new ArgumentException("Numbers must be finite."); return number; }
            if (value is byte || value is sbyte || value is short || value is ushort || value is int || value is uint || value is long || value is ulong || value is decimal)
            { decimal number = Convert.ToDecimal(value, CultureInfo.InvariantCulture); return decimal.Truncate(number) == number && number >= int.MinValue && number <= int.MaxValue ? (object)(int)number : number; }
            throw new ArgumentException("Only JSON scalars and bounded rectangular scalar arrays are supported; dictionaries, COM objects, Dates and nested rank 3 arrays are refused.");
        }

                /// <summary>Refuse tout objet arbitraire ou tableau imbriqué dans un retour censé contenir des scalaires.</summary>
                /// <param name="value">Élément scalaire retourné par la macro.</param>
                /// <param name="cells">Compteur partagé des cellules normalisées.</param>
                /// <param name="text">Compteur partagé des caractères copiés.</param>
                /// <returns>Valeur normalisée pour JSON.</returns>
        private static object JsonScalar(object value, ref int cells, ref int text)
        {
            if (value == null || value == DBNull.Value) { cells++; return null; }
            if (Marshal.IsComObject(value) || value is Array) throw new InvalidOperationException("Returned COM objects, classes and nested arrays are unsupported.");
            try { return NativeScalar(value, ref cells, ref text); }
            catch (ArgumentException ex) { throw new InvalidOperationException("The procedure returned a non-JSON or oversized scalar value.", ex); }
        }

                /// <summary>Vérifie la correspondance explicite d'un sous-type natif avec le type ByVal du paramètre.</summary>
                /// <param name="value">Valeur JSON capturée à lier.</param>
                /// <param name="type">Type scalaire déclaré par le paramètre VBA.</param>
                /// <returns>Valeur convertie au sous-type attendu par VBA.</returns>
        private static object Coerce(object value, string type)
        {
            if (Same(type, "Variant")) return value;
            if (value is Array || value == DBNull.Value || value == null) throw new ArgumentException("Array/null arguments require a ByVal Variant parameter.");
            if (Same(type, "String")) { if (value is string) return value; throw new ArgumentException("String parameters require a JSON string."); }
            if (Same(type, "Boolean")) { if (value is bool) return value; throw new ArgumentException("Boolean parameters require a JSON boolean."); }
            if (value is bool || value is string) throw new ArgumentException("Numeric parameters require a finite JSON number.");
            if (Same(type, "Double")) return Convert.ToDouble(value, CultureInfo.InvariantCulture);
            if (Same(type, "Single")) { float single = Convert.ToSingle(value, CultureInfo.InvariantCulture); if (float.IsInfinity(single)) throw new ArgumentException("Single overflow."); return single; }
            decimal number;
            try { number = Convert.ToDecimal(value, CultureInfo.InvariantCulture); }
            catch (Exception ex) when (ex is OverflowException || ex is InvalidCastException) { throw new ArgumentException("The number is outside the supported VBA numeric range.", ex); }
            if (Same(type, "Currency"))
            { if (number < -922337203685477.5808m || number > 922337203685477.5807m || decimal.Round(number, 4) != number) throw new ArgumentException("Currency requires an exact four-decimal value in range."); return new CurrencyWrapper(number); }
            if (decimal.Truncate(number) != number) throw new ArgumentException("Integral VBA parameters reject fractional values.");
            if (Same(type, "Byte") && number >= 0 && number <= 255) return (byte)number;
            if (Same(type, "Integer") && number >= short.MinValue && number <= short.MaxValue) return (short)number;
            if (Same(type, "Long") && number >= int.MinValue && number <= int.MaxValue) return (int)number;
            throw new ArgumentException("The number exceeds the explicitly declared VBA parameter range.");
        }
                /// <summary>Types de valeurs compatibles avec la liaison bornée, sans Date ni objet.</summary>
                /// <param name="type">Nom de type apparaissant dans la signature VBA.</param>
                /// <returns><see langword="true"/> si le type est accepté par le contrat JSON scalaire.</returns>
        private static bool ScalarType(string type) => new[] { "variant", "string", "boolean", "byte", "integer", "long", "single", "double", "currency" }.Contains(type.ToLowerInvariant());
                /// <summary>Identités VBA sans distinction de casse.</summary>
                /// <param name="first">Premier identifiant à comparer.</param>
                /// <param name="second">Second identifiant à comparer.</param>
                /// <returns><see langword="true"/> lorsque les deux identifiants sont égaux sans casse.</returns>
        private static bool Same(string first, string second) => string.Equals(first, second, StringComparison.OrdinalIgnoreCase);
    }
}
