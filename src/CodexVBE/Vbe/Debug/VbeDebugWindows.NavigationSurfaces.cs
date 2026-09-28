using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Web.Script.Serialization;
using System.Windows.Automation;

namespace CodexVBE
{
    /// <summary>Inspecte les surfaces de navigation visibles et leurs actions accessibles.</summary>
    internal static partial class VbeDebugWindows
    {
        /// <summary>État réellement exposé d'un nœud de navigation natif.</summary>
        internal sealed class NavigationNode
        {
            /// <summary>Jetons d’identité du nœud, de son parent, et texte et type observés.</summary>
            public string Token, ParentToken, Name, Kind, Expansion, Error;
            /// <summary>État d’activation lu pour le nœud.</summary>
            public bool Enabled;
            /// <summary>Sélection UIA directement lisible et sélection observée par MSAA.</summary>
            public bool? Selected, ObservedSelected;
            /// <summary>État MSAA brut quand le fournisseur le retourne.</summary>
            public int? MSAAState;
            /// <summary>Motif pour lequel l’action de sélection n’est pas disponible.</summary>
            public string ActionUnavailableReason;
        }

        /// <summary>Instantané d'une seule surface native ; les enfants non exposés ne sont pas inventés.</summary>
        internal sealed class NavigationSurface
        {
            /// <summary>Disponibilité d’une surface native unique et lisible.</summary>
            public bool Available;
            /// <summary>Titre, identité, erreur, fournisseur et portée du contenu retourné.</summary>
            public string Caption, Identity, Error, Provider, Coverage;
            /// <summary>Indique si les boutons de la surface sont réellement exposés.</summary>
            public bool? ButtonsExposed;
            /// <summary>État MSAA brut du conteneur éventuel.</summary>
            public int? MSAAContainerState;
            /// <summary>Nœuds accessibles de la surface, sans descendants supposés.</summary>
            public NavigationNode[] Nodes = new NavigationNode[0];
        }

        /// <summary>Accès injectable aux arbres, onglets et listes de navigation.</summary>
        internal interface INavigationSurfaceProbe
        {
            /// <summary>Lit une surface de navigation nommée.</summary>
            /// <param name="pane">Type de surface : projet, toolbox ou macros.</param>
            /// <returns>Instantané des nœuds lisibles.</returns>
            NavigationSurface Read(string pane);
            /// <summary>Exécute une action accessible sur un nœud identifiée par son jeton courant.</summary>
            /// <param name="pane">Type de surface native.</param>
            /// <param name="token">Jeton du nœud issu de l’instantané courant.</param>
            /// <param name="action">Action select, expand ou collapse.</param>
            void Act(string pane, string token, string action);
        }

                /// <summary>Lit l'explorateur de projets, la boîte à outils ou le sélecteur de macros déjà visible.</summary>
        /// <param name="request">Surface et paramètres de pagination ou de filtre.</param>
        /// <returns>Instantané paginé des seuls nœuds exposés par le fournisseur natif.</returns>
        internal static object ReadNavigationSurface(Request request) => ReadNavigationSurface(request, new NativeNavigationSurfaceProbe());

                /// <summary>Filtre un instantané complet sans utiliser le libellé comme identité de nœud.</summary>
        /// <param name="request">Surface, filtre, pagination et éventuel contrôle cible.</param>
        /// <param name="probe">Fournisseur de lecture UIA/MSAA.</param>
        /// <returns>Nœuds correspondants et version de la surface observée.</returns>
        internal static object ReadNavigationSurface(Request request, INavigationSurfaceProbe probe)
        {
            ValidateNavigationRequest(request, false);
            var state = probe.Read(request.Pane);
            ValidateNavigationSnapshot(state);
            var all = state.Nodes;
            var filtered = all.Where(n => string.IsNullOrEmpty(request.Query) || (n.Name ?? "").IndexOf(request.Query, StringComparison.OrdinalIgnoreCase) >= 0).ToArray();
            int limit = request.Limit == 0 ? 100 : request.Limit;
            return new { request.Pane, state.Available, state.Caption, state.Error,
                WindowVersion = state.Available ? NavigationRevision(state) : null,
                Nodes = filtered.Skip(request.Offset).Take(limit).ToArray(), Total = filtered.Length,
                request.Offset, Limit = limit, HasMore = request.Offset + limit < filtered.Length,
                Coverage = state.Coverage ?? "Native exposed nodes only; collapsed or inaccessible descendants are not assumed absent.",
                state.Provider, state.ButtonsExposed, state.MSAAContainerState,
                ProjectBindingVerified = false };
        }

                /// <summary>Sélectionne, développe ou replie un nœud exact après contrôle de l'ensemble de la surface.</summary>
        /// <param name="request">Surface, jeton, action et version attendue.</param>
        /// <returns>État du nœud et vérification après relecture native.</returns>
        internal static object ChangeNavigationSurface(Request request) => ChangeNavigationSurface(request, new NativeNavigationSurfaceProbe());

                /// <summary>Vérifie l'effet accessible d'une action ; une exception de livraison interdit une relance automatique.</summary>
        /// <param name="request">Action demandée avec l’empreinte de la surface inspectée.</param>
        /// <param name="probe">Fournisseur natif de lecture et d’action.</param>
        /// <returns>État avant/après, résultat observé et erreurs éventuelles.</returns>
        internal static object ChangeNavigationSurface(Request request, INavigationSurfaceProbe probe)
        {
            ValidateNavigationRequest(request, true);
            var before = probe.Read(request.Pane);
            ValidateNavigationSnapshot(before);
            if (!before.Available || !string.Equals(NavigationRevision(before), request.ExpectedWindowVersion, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The navigation surface changed or is unavailable; inspect it again.");
            var matches = before.Nodes.Where(n => n.Token == request.Control).ToArray();
            if (matches.Length != 1 || !matches[0].Enabled || !string.IsNullOrEmpty(matches[0].Error))
                throw new InvalidOperationException("The exact native node is absent, ambiguous, disabled or unreadable.");
            var target = matches[0];
            if (request.Action != "select" && (request.Pane != "project" || target.Expansion != "Expanded" && target.Expansion != "Collapsed"))
                throw new InvalidOperationException("Only expandable project tree nodes support expand/collapse.");
            if (request.Action == "select" && !target.Selected.HasValue)
                throw new InvalidOperationException(target.ActionUnavailableReason ?? "This native node exposes no selection state.");
            string beforeIdentity = before.Identity;
            bool already = NavigationDesired(target, request.Action);
            string deliveryError = null;
            if (!already)
                try { probe.Act(request.Pane, request.Control, request.Action); }
                catch (Exception) { deliveryError = "The native navigation action did not return normally."; }
            NavigationSurface after = null;
            string readError = null;
            try { after = probe.Read(request.Pane); ValidateNavigationSnapshot(after); }
            catch (Exception) { readError = "Native navigation readback failed."; }
            var observed = after == null ? new NavigationNode[0] : after.Nodes.Where(n => n.Token == request.Control).ToArray();
            bool verified = deliveryError == null && readError == null && after.Available && after.Identity == beforeIdentity &&
                observed.Length == 1 && string.IsNullOrEmpty(observed[0].Error) && NavigationDesired(observed[0], request.Action);
            return new { request.Pane, request.Control, request.Action, Applied = already ? (bool?)false : deliveryError == null ? (bool?)true : null,
                Verified = verified, VerificationPending = !verified, DeliveryError = deliveryError, ReadbackError = readError,
                Node = observed.Length == 1 ? observed[0] : null,
                WindowVersion = after != null && after.Available ? NavigationRevision(after) : null,
                NextRead = "read_navigation_surface", SourceEdited = false, PersistenceVerified = false };
        }

                /// <summary>Valide les actions de navigation et les limites de pagination.</summary>
        /// <param name="request">Requête à contrôler.</param>
        /// <param name="mutation">Exige les informations supplémentaires de mutation si vrai.</param>
        private static void ValidateNavigationRequest(Request request, bool mutation)
        {
            if (request == null || (request.Pane != "project" && request.Pane != "toolbox" && request.Pane != "macros") ||
                request.Offset < 0 || request.Offset > 100000 || request.Limit < 0 || request.Limit > 500 || (request.Query?.Length ?? 0) > 256)
                throw new ArgumentException("Pane must be project/toolbox/macros; Offset >= 0, Limit 0..500 and Query <= 256 characters.");
            if (mutation && (string.IsNullOrWhiteSpace(request.Control) || request.Control.Length > 256 ||
                string.IsNullOrWhiteSpace(request.ExpectedWindowVersion) ||
                request.Action != "select" && request.Action != "expand" && request.Action != "collapse"))
                throw new ArgumentException("Control token, ExpectedWindowVersion and Action select/expand/collapse are required.");
        }

                /// <summary>Refuse un fournisseur incomplet plutôt que présenter un arbre vide comme preuve.</summary>
        /// <param name="state">Instantané de surface à vérifier.</param>
        private static void ValidateNavigationSnapshot(NavigationSurface state)
        {
            if (state == null || state.Nodes == null || state.Nodes.Length > 4096 || state.Nodes.Any(n => n == null || string.IsNullOrEmpty(n.Token)))
                throw new InvalidOperationException("The native navigation snapshot is invalid or exceeds 4096 nodes.");
            if (state.Available && (string.IsNullOrEmpty(state.Identity) || state.Nodes.Length == 0))
                throw new InvalidOperationException("The visible navigation surface exposes no identifiable nodes.");
        }

                /// <summary>Empreinte des identités, états et erreurs de tous les nœuds observés.</summary>
        /// <param name="state">Surface native capturée.</param>
        /// <returns>Empreinte SHA-256 de l’instantané sérialisé.</returns>
        private static string NavigationRevision(NavigationSurface state)
        {
            using (var sha = SHA256.Create())
                return BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(new JavaScriptSerializer().Serialize(state)))).Replace("-", "").ToLowerInvariant();
        }

                /// <summary>Compare seulement l'état effectivement accessible correspondant à l'action.</summary>
        /// <param name="node">État lu d’un nœud.</param>
        /// <param name="action">Action dont l’effet est attendu.</param>
        /// <returns><see langword="true"/> si l’état observable correspond à l’action.</returns>
        private static bool NavigationDesired(NavigationNode node, string action) => action == "select" ? node.Selected == true :
            node.Expansion == (action == "expand" ? "Expanded" : "Collapsed");

        /// <summary>Utilise les identités UI Automation et ses patterns sans clavier ni coordonnées souris.</summary>
        private sealed class NativeNavigationSurfaceProbe : INavigationSurfaceProbe
        {
            /// <summary>Éléments UIA associés aux jetons de la dernière lecture.</summary>
            private readonly Dictionary<string, AutomationElement> elements = new Dictionary<string, AutomationElement>(StringComparer.Ordinal);
            /// <summary>Reads the live UIA parent; isolates provider detachment races during contract qualification.</summary>
            internal Func<AutomationElement, AutomationElement> ReadParent = TreeWalker.ControlViewWalker.GetParent;

            /// <summary>Recherche et lit la surface native unique correspondant au nom demandé.</summary>
            /// <param name="pane">Type de surface : projet, toolbox ou macros.</param>
            /// <returns>Instantané UI Automation ou motif d’indisponibilité.</returns>
            public NavigationSurface Read(string pane)
            {
                elements.Clear();
                var result = new NavigationSurface();
                try
                {
                    IntPtr root = FindVbeRoot();
                    if (root == IntPtr.Zero) { result.Error = "The VBE window is unavailable."; return result; }
                    GetWindowThreadProcessId(root, out uint rootOwner);
                    var windows = new List<IntPtr>();
                    EnumChildWindows(root, (h, p) => {
                        if (!IsWindowVisible(h)) return true;
                        string caption = WindowText(h), kind = ClassName(h);
                        if (pane == "project" && kind == "SysTreeView32") windows.Add(h);
                        else if (pane == "toolbox" && (kind == "VbaWindow" || kind.StartsWith("F3 MinFrame ", StringComparison.Ordinal)) &&
                            (caption == "Toolbox" || caption == "Boîte à outils")) windows.Add(h);
                        return true;
                    }, IntPtr.Zero);
                    if (pane == "macros")
                    {
                        IntPtr dialog = FindDialog("Macros", "Macro");
                        if (dialog != IntPtr.Zero) windows.Add(dialog);
                    }
                    if (pane == "toolbox")
                    {
                        EnumWindows((h, p) => {
                            GetWindowThreadProcessId(h, out uint pid);
                            if (pid == rootOwner && rootOwner != 0 && IsWindowVisible(h) && ClassName(h).StartsWith("F3 MinFrame ", StringComparison.Ordinal) &&
                                (WindowText(h) == "Toolbox" || WindowText(h) == "Boîte à outils") && !windows.Contains(h)) windows.Add(h);
                            return true;
                        }, IntPtr.Zero);
                        windows.RemoveAll(h => { GetWindowThreadProcessId(h, out uint pid); return pid != rootOwner || rootOwner == 0; });
                    }
                    if (windows.Count != 1) { result.Error = "The exact native navigation surface is absent or ambiguous."; return result; }
                    IntPtr handle = windows[0];
                    if (pane == "toolbox")
                    {
                        // UIA can project the same F3 page through several wrappers without
                        // exposing selection. Prefer the canonical MSAA grouping provider.
                        var pages = ReadNativeToolboxPages(handle, rootOwner);
                        if (pages.Available) return pages;
                    }
                    try
                    {
                        var surface = AutomationElement.FromHandle(handle);
                        var types = pane == "project" ? new[] { ControlType.TreeItem } : pane == "toolbox" ?
                            new[] { ControlType.TabItem, ControlType.Button, ControlType.ListItem } : new[] { ControlType.ListItem };
                        var condition = new OrCondition(types.Concat(new[] { ControlType.TreeItem }).Distinct()
                            .Select(t => (Condition)new PropertyCondition(AutomationElement.ControlTypeProperty, t)).Concat(new[] { Condition.FalseCondition }).ToArray());
                        var nodes = surface.FindAll(TreeScope.Descendants, condition);
                        if (nodes.Count > 4096) throw new InvalidOperationException("Too many native nodes.");
                        var list = new List<NavigationNode>();
                        foreach (AutomationElement element in nodes)
                        {
                            string token = Token(element);
                            var node = new NavigationNode { Token = token, Name = element.Current.Name,
                                Kind = element.Current.ControlType.ProgrammaticName, Enabled = element.Current.IsEnabled };
                            if (elements.ContainsKey(token)) throw new InvalidOperationException("Duplicate native node identity.");
                            elements.Add(token, element);
                            try
                            {
                                if (element.TryGetCurrentPattern(SelectionItemPattern.Pattern, out object selection))
                                    node.Selected = ((SelectionItemPattern)selection).Current.IsSelected;
                                if (element.TryGetCurrentPattern(ExpandCollapsePattern.Pattern, out object expansion))
                                    node.Expansion = ((ExpandCollapsePattern)expansion).Current.ExpandCollapseState.ToString();
                                var parent = ReadParent(element);
                                if (parent != null && parent.Current.ControlType == ControlType.TreeItem) node.ParentToken = Token(parent);
                            }
                            catch (Exception) { node.Error = "The native node state could not be read."; }
                            list.Add(node);
                        }
                        result.Caption = WindowText(handle);
                        result.Identity = handle.ToInt64().ToString(System.Globalization.CultureInfo.InvariantCulture);
                        result.Nodes = list.ToArray();
                        result.Available = list.Count > 0;
                        if (!result.Available && pane == "toolbox") return ReadNativeToolboxPages(handle, rootOwner);
                        if (!result.Available) result.Error = "No native nodes are exposed; the surface is not certified empty.";
                        }
                    catch (Exception) when (pane == "toolbox") { return ReadNativeToolboxPages(handle, rootOwner); }
                }
                catch (Exception) { result.Available = false; result.Error = "The native navigation surface is inaccessible."; }
                return result;
            }

            /// <summary>Utilise le pattern UIA du jeton courant pour sélectionner ou modifier son expansion.</summary>
            /// <param name="pane">Type de surface où le jeton a été émis.</param>
            /// <param name="token">Identité UIA du nœud lu lors du dernier appel à <see cref="Read"/>.</param>
            /// <param name="action">Action select, expand ou collapse.</param>
            public void Act(string pane, string token, string action)
            {
                if (!elements.TryGetValue(token, out AutomationElement element)) throw new InvalidOperationException("Node identity expired.");
                if (action == "select")
                    ((SelectionItemPattern)element.GetCurrentPattern(SelectionItemPattern.Pattern)).Select();
                else
                {
                    var pattern = (ExpandCollapsePattern)element.GetCurrentPattern(ExpandCollapsePattern.Pattern);
                    if (action == "expand") pattern.Expand(); else pattern.Collapse();
                }
                PauseNative(75);
            }

                        /// <summary>Identité de session du fournisseur UIA ; aucun nom seul n'identifie un nœud.</summary>
            /// <param name="element">Élément UI Automation à identifier.</param>
            /// <returns>Jeton construit à partir de l’identifiant d’exécution UIA.</returns>
            private static string Token(AutomationElement element) => string.Join(".", element.GetRuntimeId());
        }
    }
}
