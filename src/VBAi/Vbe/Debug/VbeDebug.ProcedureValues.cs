using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;

namespace VBAi
{
    /// <summary>Prépare et suit l’invocation de procédures VBA avec valeurs JSON et retour sérialisable.</summary>
    internal sealed partial class VbeDebug
    {
        /// <summary>Transport natif pour une seule invocation avec valeurs; aucune expression VBA ni helper n'est injecté.</summary>
        internal interface IProcedureValuesHost
        {
                        /// <summary>Résout un classeur ouvert par identité COM du VBProject et vérifie son PID et son chemin.</summary>
                        /// <param name="project">Projet VBIDE dont le classeur propriétaire doit être trouvé.</param>
                        /// <param name="expectedHostPath">Chemin absolu du classeur attendu lors de l’inspection.</param>
                        /// <returns>Cible détenue qui associe l’application Excel et le classeur ouvert correspondant.</returns>
            object ResolveTarget(object project, string expectedHostPath);
                        /// <summary>Invoque la procédure standard résolue exactement une fois.</summary>
                        /// <param name="target">Cible Excel détenue retournée par <see cref="ResolveTarget"/>.</param>
                        /// <param name="module">Nom du module standard relu.</param>
                        /// <param name="procedure">Nom de la procédure relue.</param>
                        /// <param name="arguments">Arguments convertis dans l’ordre d’appel.</param>
                        /// <returns>Valeur native retournée par Excel.Run.</returns>
            object Invoke(object target, string module, string procedure, object[] arguments);
        }

        /// <summary>Transport Excel.Application.Run dans le processus courant; autres hôtes non pris en charge.</summary>
        internal sealed class NativeProcedureValuesHost : IProcedureValuesHost
        {
            /// <summary>Reads the process identity used to restrict the transport to an Excel host.</summary>
            internal Func<string> ReadProcessName = CurrentProcessName;
            /// <summary>Resolves only the application owned by the supplied process; replaceable for contract hosts.</summary>
            internal Func<int, Func<object>, object> ResolveApplication = ExcelOwnedApplication.Resolve;
            /// <summary>Reads an already registered application without starting a process.</summary>
            internal Func<string, object> ReadActiveApplication = Marshal.GetActiveObject;
            /// <summary>Performs the current process name operation for NativeProcedureValuesHost.</summary>
            /// <returns>The result produced by this operation.</returns>
            private static string CurrentProcessName()
            { using (var process = System.Diagnostics.Process.GetCurrentProcess()) return process.ProcessName; }
            /// <summary>Application/classeur COM détenus pour le seul appel préparé.</summary>
            private sealed class OwnedTarget
            {
                /// <summary>Application Excel du processus courant.</summary>
                internal object Application, Workbook;
                /// <summary>Chemin absolu validé du classeur associé.</summary>
                internal string Path;
            }
                        /// <summary>Obtient Excel déjà ouvert dans le PID de l'add-in et associe un seul classeur au VBProject.</summary>
                        /// <param name="project">Projet VBIDE dont l’identité COM est comparée.</param>
                        /// <param name="expectedHostPath">Chemin complet attendu du classeur enregistré.</param>
                        /// <returns>Cible détenue associant Excel, le classeur et le chemin vérifié.</returns>
            public object ResolveTarget(object project, string expectedHostPath)
            {
                using (var process = System.Diagnostics.Process.GetCurrentProcess())
                {
                    if (!string.Equals(ReadProcessName(), "EXCEL", StringComparison.OrdinalIgnoreCase))
                        throw new InvalidOperationException("Returned procedure values currently require the in-process Excel host.");
                    dynamic application = ResolveApplication(process.Id, () => ReadActiveApplication("Excel.Application"));
                    uint owner; VbeDebugWindows.GetWindowThreadProcessId(new IntPtr(Convert.ToInt64(application.Hwnd)), out owner);
                    if (owner != (uint)process.Id) throw new InvalidOperationException("The Excel application belongs to another PID.");
                    object match = null; int count = 0;
                    foreach (dynamic workbook in application.Workbooks)
                    {
                        if (++count > 1000) throw new InvalidOperationException("Unexpected open workbook count.");
                        if (!SameComIdentity(project, (object)workbook.VBProject)) continue;
                        if (match != null) throw new InvalidOperationException("Multiple workbooks share the selected project identity.");
                        if (!SameValuesPath((string)workbook.FullName, expectedHostPath) || string.IsNullOrWhiteSpace((string)workbook.Path))
                            throw new InvalidOperationException("The owned workbook path changed or has never been saved.");
                        match = (object)workbook;
                    }
                    if (match == null) throw new InvalidOperationException("No owned open workbook shares this VBProject COM identity.");
                    return new OwnedTarget { Application = (object)application, Workbook = match, Path = expectedHostPath };
                }
            }
                        /// <summary>Envoie les paramètres par position à Excel.Run; les tableaux sont marshalisés comme SAFEARRAY de VARIANT.</summary>
                        /// <param name="target">Cible détenue obtenue avant l’appel.</param>
                        /// <param name="module">Nom du module standard cible.</param>
                        /// <param name="procedure">Nom de la procédure à exécuter.</param>
                        /// <param name="arguments">Arguments validés par la liaison de signature.</param>
                        /// <returns>Objet natif que la procédure retourne, avant sa normalisation JSON.</returns>
            public object Invoke(object target, string module, string procedure, object[] arguments)
            {
                var owned = target as OwnedTarget;
                if (owned == null) throw new InvalidOperationException("An owned Excel target is required.");
                if (!SameValuesPath((string)((dynamic)owned.Workbook).FullName, owned.Path))
                    throw new InvalidOperationException("The workbook path changed before invocation.");
                var invokeArguments = new object[arguments.Length + 1];
                invokeArguments[0] = "'" + owned.Path.Replace("'", "''") + "'!" + module + "." + procedure;
                Array.Copy(arguments, 0, invokeArguments, 1, arguments.Length);
                return owned.Application.GetType().InvokeMember("Run", BindingFlags.InvokeMethod | BindingFlags.OptionalParamBinding,
                    null, owned.Application, invokeArguments, System.Globalization.CultureInfo.InvariantCulture);
            }
                        /// <summary>Compare IUnknown sans utiliser les noms de projet ou de classeur.</summary>
                        /// <param name="first">Première référence COM.</param>
                        /// <param name="second">Seconde référence COM.</param>
                        /// <returns><see langword="true"/> si les deux références désignent le même IUnknown.</returns>
            internal static bool SameComIdentity(object first, object second)
            {
                if (first == null || second == null || !Marshal.IsComObject(first) || !Marshal.IsComObject(second)) return false;
                IntPtr a = IntPtr.Zero, b = IntPtr.Zero;
                try { a = Marshal.GetIUnknownForObject(first); b = Marshal.GetIUnknownForObject(second); return a == b; }
                finally { if (a != IntPtr.Zero) Marshal.Release(a); if (b != IntPtr.Zero) Marshal.Release(b); }
            }
        }

        /// <summary>Transport remplaçable par une sonde de tests; le transport natif ne démarre aucun hôte.</summary>
        internal IProcedureValuesHost ProcedureValuesHost = new NativeProcedureValuesHost();
        /// <summary>Identifiants des appels à valeurs et état de leur unique tentative native.</summary>
        private readonly Dictionary<string, bool> procedureValueInvocations = new Dictionary<string, bool>();

        /// <summary>Plan sans mutation de code, limité à la signature et aux valeurs déjà copiées.</summary>
        private sealed class ValuesCall
        {
            /// <summary>Module de la procédure relue.</summary>
            internal string Module, Procedure, Identity;
            /// <summary>Cible Excel associée au projet sélectionné.</summary>
            internal object Target;
            /// <summary>Arguments convertis selon la signature active.</summary>
            internal object[] Arguments;
        }

        /// <summary>Planifie une invocation Excel.Run avec valeurs scalaires/tableaux JSON et capture du retour natif.</summary>
        /// <param name="request">Projet, module standard, procédure, SHA, ExpectedHostPath exact, ExpectedMode=2 et Arguments.</param>
        /// <returns>Identifiant à lire via procedure_values_status; aucun succès d'exécution n'est présumé à la mise en file.</returns>
        public object RunProcedureValues(Request request)
        {
            if (request == null) throw new ArgumentNullException(nameof(request));
            var context = SynchronizationContext.Current;
            if (context == null) throw new InvalidOperationException("The VBE UI context is unavailable.");
            if (procedureOperations.Any(x => x.State == "Queued" || x.State == "Delivering"))
                throw new InvalidOperationException("A procedure call is pending; inspect its status instead of retrying.");
            var captured = new Request { Project = request.Project, Module = request.Module, Procedure = request.Procedure,
                ExpectedSha256 = request.ExpectedSha256, ExpectedHostPath = request.ExpectedHostPath, ExpectedMode = request.ExpectedMode,
                Arguments = VbaProcedureValues.Capture(request.Arguments),
                ArgumentNames = request.ArgumentNames == null ? null : (string[])request.ArgumentNames.Clone() };
            var prepared = PrepareProcedureValues(captured);
            var operation = new ProcedureOperation { Id = Guid.NewGuid().ToString("N"), Project = captured.Project, Module = prepared.Module,
                Procedure = prepared.Procedure, Command = prepared.Identity, State = "Queued" };
            procedureOperations.Add(operation); procedureValueInvocations.Add(operation.Id, false);
            if (procedureOperations.Count > 20)
            { procedureValueInvocations.Remove(procedureOperations[0].Id); procedureOperations.RemoveAt(0); }
            foreach (string expired in procedureValueInvocations.Keys.Where(id => !procedureOperations.Any(item => item.Id == id)).ToArray())
                procedureValueInvocations.Remove(expired);
            try
            {
                context.Post(_ => {
                    try
                    {
                        var live = PrepareProcedureValues(captured);
                        if (live.Identity != prepared.Identity) throw new InvalidOperationException("Procedure identity changed before execution.");
                        operation.State = "Delivering"; procedureValueInvocations[operation.Id] = true;
                        object returned = ProcedureValuesHost.Invoke(live.Target, live.Module, live.Procedure, live.Arguments);
                        operation.State = "ReturnReceived";
                        operation.Output = VbaProcedureValues.NormalizeReturn(returned);
                        operation.State = "Returned";
                    }
                    catch (Exception ex)
                    {
                        operation.Error = ex.Message;
                        operation.State = operation.State == "ReturnReceived" ? "ReturnRejected" : "Failed";
                    }
                }, null);
            }
            catch (Exception ex) { operation.State = "Failed"; operation.Error = ex.Message; }
            return ProcedureValuesResult(operation);
        }

                /// <summary>Lit le résultat d'une invocation sans réévaluer la procédure.</summary>
                /// <param name="request">Requête de statut portant le projet et l’identifiant de l’invocation.</param>
                /// <returns>État courant de l’unique invocation et valeur normalisée si elle a été retournée.</returns>
        public object ProcedureValuesStatus(Request request)
        {
            if (request == null) throw new ArgumentNullException(nameof(request));
            var operation = procedureOperations.FirstOrDefault(x => x.Id == request.Query && procedureValueInvocations.ContainsKey(x.Id) &&
                string.Equals(x.Project, request.Project, StringComparison.OrdinalIgnoreCase));
            if (operation == null) throw new InvalidOperationException("Unknown values operation in this project/session.");
            return ProcedureValuesResult(operation);
        }

                /// <summary>Relit mode, identité, chemin complet, signature et SHA avant chaque tentative native.</summary>
                /// <param name="request">Requête déjà capturée avec chemin, mode et SHA attendus.</param>
                /// <returns>Plan validé contenant l’identité, la cible détenue et les arguments convertis.</returns>
        private ValuesCall PrepareProcedureValues(Request request)
        {
            if (request.ExpectedMode != 2) throw new ArgumentException("ExpectedMode=2 is required.");
            ValidateProcedureIdentifier(request.Module); ValidateProcedureIdentifier(request.Procedure);
            string expectedPath = RequireValuesPath(request.ExpectedHostPath);
            dynamic project = GetProject(request.Project);
            if ((int)project.Mode != 2) throw new InvalidOperationException("The project is no longer in design mode.");
            if (!SameValuesPath((string)project.FileName, expectedPath)) throw new InvalidOperationException("The selected project's host path changed.");
            var components = new List<object>();
            foreach (dynamic candidate in project.VBComponents)
                if (string.Equals((string)candidate.Name, request.Module, StringComparison.OrdinalIgnoreCase)) components.Add((object)candidate);
            if (components.Count != 1 || (int)((dynamic)components[0]).Type != 1)
                throw new InvalidOperationException("One unique standard module is required.");
            dynamic component = components[0], codeModule = component.CodeModule; int count = (int)codeModule.CountOfLines;
            string source = count == 0 ? "" : (string)codeModule.Lines[1, count];
            if (string.IsNullOrWhiteSpace(request.ExpectedSha256) || !string.Equals(Hash(source), request.ExpectedSha256, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The source changed since inspection.");
            int body = (int)codeModule.ProcBodyLine[request.Procedure, 0];
            if (body < 1 || body > count) throw new InvalidOperationException("The exact procedure body is absent.");
            object[] bound = VbaProcedureValues.Bind(source, body, request.Procedure, request.Arguments, request.ArgumentNames);
            string module = (string)component.Name;
            object target = ProcedureValuesHost.ResolveTarget((object)project, expectedPath);
            int finalCount = (int)codeModule.CountOfLines;
            string finalSource = finalCount == 0 ? "" : (string)codeModule.Lines[1, finalCount];
            if ((int)project.Mode != 2 || !SameValuesPath((string)project.FileName, expectedPath) ||
                !string.Equals(Hash(finalSource), request.ExpectedSha256, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("Source, mode or host path changed while resolving the owned invocation target.");
            return new ValuesCall { Module = module, Procedure = request.Procedure, Target = target, Arguments = bound,
                Identity = expectedPath + "!" + module + "." + request.Procedure + ":" + request.ExpectedSha256 };
        }

                /// <summary>Décrit la preuve du retour, distincte de la correction de la macro et de ses effets de bord.</summary>
                /// <param name="operation">État suivi de l’invocation.</param>
                /// <returns>Résultat sérialisable avec statut, sortie, erreur et limites de vérification.</returns>
        private object ProcedureValuesResult(ProcedureOperation operation) => new { operation.Project, operation.Module, operation.Procedure,
            Query = operation.Id, operation.State, operation.Output, operation.Error,
            Pending = operation.State == "Queued" || operation.State == "Delivering", InvocationInvoked = procedureValueInvocations[operation.Id],
            ReturnValueVerified = operation.State == "Returned", RuntimeSuccessVerified = false,
            Uncertain = procedureValueInvocations[operation.Id] && operation.State == "Failed", Transport = "OwnedExcelApplicationRun",
            NextRead = "procedure_values_status, debug_state, debug_dialog",
            Limit = "Excel owned-process Application.Run only. Fixed ByVal scalar/Variant parameters, positional/named binding; ParamArray Variant accepts positional values only and no Optional prefix, up to 30 total arguments. Rectangular zero-based JSON inputs rank 1/2; scalar-array returns preserve native bounds. No ByRef mutation contract, class/object/Date values or injected helper. Native calls can block on modal/runtime code. One invocation; never retry automatically." };

                /// <summary>Exige un chemin Windows absolu de fichier macro Excel déjà associé au projet.</summary>
                /// <param name="path">Chemin reçu dans la requête.</param>
                /// <returns>Chemin absolu normalisé vers un format de classeur macro pris en charge.</returns>
        private static string RequireValuesPath(string path)
        {
            if (string.IsNullOrWhiteSpace(path) || !System.Text.RegularExpressions.Regex.IsMatch(path, @"^(?:[A-Za-z]:[\\/]|\\\\[^\\/]+[\\/][^\\/]+[\\/])"))
                throw new ArgumentException("ExpectedHostPath must be the inspected fully qualified Excel workbook path.");
            if (path.IndexOfAny(new[] { '\r', '\n', '\0', '!', '[', ']' }) >= 0)
                throw new ArgumentException("The host path contains unsupported macro-qualification characters.");
            string full = Path.GetFullPath(path);
            if (!new[] { ".xlsm", ".xlsb", ".xlam", ".xltm", ".xls" }.Contains(Path.GetExtension(full).ToLowerInvariant()))
                throw new ArgumentException("A saved macro-enabled Excel workbook/add-in is required.");
            return full;
        }
                /// <summary>Compare un chemin natif au chemin absolu attendu.</summary>
                /// <param name="first">Chemin lu depuis le projet ou le classeur.</param>
                /// <param name="expected">Chemin absolu attendu.</param>
                /// <returns><see langword="true"/> si le chemin natif est absolu et identique sans distinction de casse.</returns>
        private static bool SameValuesPath(string first, string expected) => !string.IsNullOrWhiteSpace(first) && Path.IsPathRooted(first) &&
            string.Equals(Path.GetFullPath(first), expected, StringComparison.OrdinalIgnoreCase);
    }
}
