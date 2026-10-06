using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace VBAi
{

    /// <summary>Prépare, transmet et suit les appels scalaires aux procédures publiques du VBE.</summary>
    internal sealed partial class VbeDebug
    {

        /// <summary>Appels récents conservés dans cette session, sans autorité de relancement.</summary>
        private readonly List<ProcedureOperation> procedureOperations = new List<ProcedureOperation>();

        /// <summary>Frontière native asynchrone injectable pour les essais locaux.</summary>
        internal Func<string, Task<object>> ExecuteProcedureCall = command => Task.Run(() => VbeDebugWindows.ExecuteImmediate(command));

        /// <summary>Ouverture du volet Exécution, injectable pour les tests sans hôte.</summary>
        internal Action ShowProcedureImmediate;

        /// <summary>Planifie un appel paramétré, puis relit code et mode sur le thread VBE avant transmission.</summary>
        /// <param name="request">Projet, module standard, procédure publique, SHA et arguments scalaires.</param>
        /// <returns>Identifiant à transmettre à procedure_run_status.</returns>
        public object RunProcedure(Request request)
        {
            if (request == null) throw new ArgumentNullException(nameof(request));
            var context = SynchronizationContext.Current;
            if (context == null) throw new InvalidOperationException("The VBE UI context is unavailable.");
            if (procedureOperations.Any(x => x.State == "Queued" || x.State == "Delivering"))
                throw new InvalidOperationException("A procedure call is still pending; inspect its status instead of retrying.");
            var captured = new Request { Project = request.Project, Module = request.Module, Procedure = request.Procedure,
                ExpectedSha256 = request.ExpectedSha256, ExpectedMode = request.ExpectedMode,
                Arguments = request.Arguments == null ? null : (object[])request.Arguments.Clone(),
                ArgumentNames = request.ArgumentNames == null ? null : (string[])request.ArgumentNames.Clone() };
            string command = PrepareProcedureCall(captured);
            var operation = new ProcedureOperation { Id = Guid.NewGuid().ToString("N"), Project = captured.Project,
                Module = captured.Module, Procedure = captured.Procedure, State = "Queued", Command = command };
            procedureOperations.Add(operation);
            if (procedureOperations.Count > 20) procedureOperations.RemoveAt(0);
            try
            {
                context.Post(async _ => {
                    try
                    {
                        if (PrepareProcedureCall(captured) != command)
                            throw new InvalidOperationException("Procedure identity changed before execution.");
                        if (ShowProcedureImmediate != null) ShowProcedureImmediate();
                        else OpenDebugPane("immediate", new VbeEditorWindows((object)vbe));
                        operation.State = "Delivering";
                        // Native messages are processed by the UI while this worker observes their echo.
                        operation.Output = await ExecuteProcedureCall(command);
                        operation.State = "Delivered";
                    }
                    catch (Exception ex) { operation.Error = ex.Message; operation.State = "Failed"; }
                }, null);
            }
            catch (Exception ex) { operation.Error = ex.Message; operation.State = "Failed"; }
            return ProcedureResult(operation);
        }

        /// <summary>Lit le suivi d’un appel sans réexécuter de code.</summary>
        /// <param name="request">Projet exact et identifiant Query retourné lors de la planification.</param>
        /// <returns>État de transmission, sortie native ou erreur.</returns>
        public object ProcedureRunStatus(Request request)
        {
            var operation = procedureOperations.FirstOrDefault(x => x.Id == request.Query &&
                string.Equals(x.Project, request.Project, StringComparison.OrdinalIgnoreCase));
            if (operation == null) throw new InvalidOperationException("Unknown procedure operation in this project/session.");
            return ProcedureResult(operation);
        }

        /// <summary>Construit uniquement un appel à une déclaration publique du module vivant vérifié.</summary>
        /// <param name="request">Identité, version et arguments de l’appel.</param>
        /// <returns>Instruction VBA monoligne sans code fourni par le modèle.</returns>
        internal string PrepareProcedureCall(Request request)
        {
            if (request.ExpectedMode != 2 || request.Arguments == null || request.Arguments.Length > 30)
                throw new ArgumentException("ExpectedMode=2 and at most 30 scalar Arguments are required.");
            ValidateProcedureIdentifier(request.Procedure);
            dynamic project = GetProject(request.Project);
            if ((int)project.Mode != 2) throw new InvalidOperationException("The selected project is no longer in design mode.");
            string projectName = (string)project.Name;
            ValidateProcedureIdentifier(projectName);
            int sameName = 0;
            foreach (dynamic candidate in vbe.VBProjects)
                if (string.Equals((string)candidate.Name, projectName, StringComparison.OrdinalIgnoreCase)) sameName++;
            if (sameName != 1) throw new InvalidOperationException("A fully qualified Immediate call requires a unique project name.");
            dynamic component = null;
            foreach (dynamic candidate in project.VBComponents)
                if (string.Equals((string)candidate.Name, request.Module, StringComparison.OrdinalIgnoreCase)) { component = candidate; break; }
            if (component == null || (int)component.Type != 1)
                throw new InvalidOperationException("A public procedure in a standard module is required.");
            string moduleName = (string)component.Name;
            ValidateProcedureIdentifier(moduleName);
            dynamic module = component.CodeModule;
            int count = (int)module.CountOfLines;
            string code = count == 0 ? "" : (string)module.Lines[1, count];
            if (string.IsNullOrWhiteSpace(request.ExpectedSha256) || !string.Equals(Hash(code), request.ExpectedSha256, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The module changed since it was read.");
            int body = (int)module.ProcBodyLine[request.Procedure, 0];
            if (body < 1 || body > count) throw new InvalidOperationException("The procedure declaration is not in the module.");
            string declaration = (string)module.Lines[body, 1];
            var match = Regex.Match(declaration, @"^\s*(?:(?:Public|Static)\s+)*(Sub|Function)\s+(" + Regex.Escape(request.Procedure) + @")(?![\p{L}\p{N}_])", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            if (!match.Success) throw new InvalidOperationException("Only public Sub and Function declarations are callable; private, property and external declarations are refused.");
            string[] literals = request.Arguments.Select(ProcedureLiteral).ToArray();
            if (request.ArgumentNames != null)
            {
                if (request.ArgumentNames.Length != literals.Length || request.ArgumentNames.Distinct(StringComparer.OrdinalIgnoreCase).Count() != literals.Length)
                    throw new ArgumentException("ArgumentNames must match Arguments and contain distinct parameter names.");
                var signature = VbaDeclarationIndex.Statements(code).First(s => s.Count > 0 && s[0].Line == body);
                int end = signature.Last().Line;
                var declared = VbaDeclarationIndex.Read(code).Where(d => d.Kind == "Parameter" && d.Line >= body && d.Line <= end &&
                    string.Equals(d.Scope, request.Procedure, StringComparison.OrdinalIgnoreCase)).ToArray();
                if (declared.Any(d => d.Conditional))
                    throw new InvalidOperationException("Named calls to conditional signatures require a resolved binding plan.");
                if (signature.Any(token => token.Text.Equals("ParamArray", StringComparison.OrdinalIgnoreCase)))
                    throw new InvalidOperationException("Named calls to ParamArray signatures are not supported.");
                for (int i = 0; i < literals.Length; i++)
                {
                    ValidateProcedureIdentifier(request.ArgumentNames[i]);
                    var parameter = declared.SingleOrDefault(d => string.Equals(d.Name, request.ArgumentNames[i], StringComparison.OrdinalIgnoreCase));
                    if (parameter == null) throw new ArgumentException("Unknown parameter in the live signature: " + request.ArgumentNames[i]);
                    literals[i] = parameter.Name + ":=" + literals[i];
                }
            }
            string arguments = string.Join(", ", literals);
            string call = projectName + "." + moduleName + "." + match.Groups[2].Value + "(" + arguments + ")";
            string command = match.Groups[1].Value.Equals("Function", StringComparison.OrdinalIgnoreCase) ? "? " + call : "Call " + call;
            if (command.Length > 2048) throw new ArgumentException("The complete procedure call exceeds the native 2048-character limit.");
            return command;
        }

        /// <summary>Encode un scalaire JSON en littéral VBA sans expression exécutable arbitraire.</summary>
        /// <param name="value">Scalaire à convertir.</param>
        /// <returns>Littéral invariant, null JSON devenant Null VBA.</returns>
        internal static string ProcedureLiteral(object value)
        {
            if (value == null) return "Null";
            if (value is string text)
            {
                if (text.Any(char.IsControl)) throw new ArgumentException("Procedure strings must contain printable characters only.");
                return "\"" + text.Replace("\"", "\"\"") + "\"";
            }
            if (value is bool flag) return flag ? "True" : "False";
            if (value is byte || value is sbyte || value is short || value is ushort || value is int || value is uint || value is long || value is ulong || value is decimal)
                return Convert.ToString(value, CultureInfo.InvariantCulture);
            if (value is double || value is float)
            {
                double number = Convert.ToDouble(value, CultureInfo.InvariantCulture);
                if (!double.IsNaN(number) && !double.IsInfinity(number)) return number.ToString("R", CultureInfo.InvariantCulture);
            }
            throw new ArgumentException("Procedure arguments must be strings, finite numbers, booleans or null.");
        }

        /// <summary>Refuse toute syntaxe autre qu’un identifiant VBA qualifiable.</summary>
        /// <param name="name">Nom de projet, module ou procédure à vérifier.</param>
        private static void ValidateProcedureIdentifier(string name)
        {
            if (!Regex.IsMatch(name ?? "", @"^\p{L}[\p{L}\p{N}_]{0,254}$")) throw new ArgumentException("An exact VBA identifier is required.");
        }

        /// <summary>Sérialise le suivi sans assimiler transmission et succès runtime.</summary>
        /// <param name="operation">Opération d’appel enregistrée en mémoire.</param>
        /// <returns>État de transmission, sortie observée et limites de vérification runtime.</returns>
        private static object ProcedureResult(ProcedureOperation operation) => new { operation.Project, operation.Module,
            operation.Procedure, Query = operation.Id, operation.State, operation.Command, operation.Output, operation.Error,
            Pending = operation.State == "Queued" || operation.State == "Delivering", RuntimeSuccessVerified = false,
            NextRead = "procedure_run_status, debug_state, debug_dialog",
            Limit = "Scalar arguments only; optional ArgumentNames must match the inspected live signature. No named ParamArray/conditional calls, object/array arguments or returned COM values. Runtime errors and modal code can outlive delivery; do not retry automatically." };

        /// <summary>État borné en mémoire d’une seule tentative d’appel.</summary>
        private sealed class ProcedureOperation
        {

            /// <summary>Identifiant de suivi et identité de la cible.</summary>
            public string Id, Project, Module, Procedure, State, Command, Error;

            /// <summary>Sortie de transmission ou résultat exposé par l’interface native.</summary>
            public object Output;
        }
    }
}
