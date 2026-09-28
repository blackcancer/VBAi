using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Automation;

namespace CodexVBE
{
    /// <summary>Lit et sélectionne les éléments de l’Explorateur d’objets par UI Automation native.</summary>
internal static partial class VbeDebugWindows
    {
        /// <summary>Native UIA element acquisition, isolated without replacing browser orchestration.</summary>
        internal static Func<IntPtr, AutomationElement> ObjectBrowserElement = AutomationElement.FromHandle;
        // Accessibility must run on a worker thread, never the VBE dispatcher.
        /// <summary>Capture les listes, sélections et descriptions accessibles de l’Explorateur d’objets.</summary>
        /// <returns>État UI Automation disponible, ou motif d’indisponibilité.</returns>
internal static object ReadObjectBrowser()
        {
            IntPtr root = FindVbeRoot();
            if (root == IntPtr.Zero) return new { Available = false, Reason = "VBE window unavailable." };
            IntPtr pane = FindPane(ChildWindows(root), "Object Browser", "Explorateur d'objets");
            if (pane == IntPtr.Zero || !IsWindowVisible(pane))
                return new { Available = false, Reason = "No visible supported Object Browser window was identified." };
            var controls = ObjectBrowserControls(pane);
            var selections = new List<object>();
            var descriptions = new List<object>();
            foreach (AutomationElement control in controls)
            {
                try
                {
                    string name = control.Current.Name;
                    string kind = control.Current.ControlType.ProgrammaticName;
                    if (control.Current.ControlType == ControlType.Document)
                    {
                        if (control.TryGetCurrentPattern(TextPattern.Pattern, out object text))
                        {
                            string value = ((TextPattern)text).DocumentRange.GetText(16385);
                            descriptions.Add(new { Name = name, Text = value.Substring(0, Math.Min(value.Length, 16384)).Trim(), Truncated = value.Length > 16384 });
                        }
                        else descriptions.Add(new { Name = name, Error = "Text pattern unavailable." });
                        continue;
                    }
                    string valueText = null;
                    if (control.TryGetCurrentPattern(ValuePattern.Pattern, out object valuePattern))
                        valueText = ((ValuePattern)valuePattern).Current.Value;
                    bool readable = control.TryGetCurrentPattern(SelectionPattern.Pattern, out object selection);
                    string[] selected = readable ? ((SelectionPattern)selection).Current.GetSelection()
                        .Select(item => item.Current.Name).ToArray() : null;
                    selections.Add(new { Name = name, Type = kind, Enabled = control.Current.IsEnabled, Handle = control.Current.NativeWindowHandle, NativeEnabled = ObjectBrowserEnabled(new IntPtr(control.Current.NativeWindowHandle)), NativeClass = ClassName(new IntPtr(control.Current.NativeWindowHandle)), Value = valueText, SelectionReadable = readable, Selected = selected });
                }
                catch (Exception ex) { selections.Add(new { Error = ex.Message }); }
            }
            return new { Available = true, Caption = WindowText(pane), Selections = selections,
                Descriptions = descriptions, Source = "Native VBE UI Automation", SemanticResolutionVerified = false };
        }

        /// <summary>Énumère les seuls HWND exposant une liste, une combo ou une zone Document reconnue.</summary>
        /// <param name="pane">Handle du volet Explorateur d’objets.</param>
        /// <returns>Éléments UI Automation des contrôles pris en charge.</returns>
private static List<AutomationElement> ObjectBrowserControls(IntPtr pane)
        {
            // Enumerate HWNDs rather than every accessible list item: COM libraries
            // can expose thousands of classes and members beneath these controls.
            var controls = new List<AutomationElement>();
            EnumChildWindows(pane, (handle, ignored) =>
            {
                try
                {
                    AutomationElement element = ObjectBrowserElement(handle);
                    ControlType kind = element.Current.ControlType;
                    if (kind == ControlType.List || kind == ControlType.ComboBox || kind == ControlType.Document)
                        controls.Add(element);
                }
                catch { /* Individual inaccessible HWNDs do not identify a supported control. */ }
                return true;
            }, IntPtr.Zero);
            return controls;
        }

        /// <summary>Obtient le parent natif d’un contrôle de l’Explorateur d’objets.</summary>
        /// <param name="handle">Handle du contrôle.</param>
        /// <returns>Handle parent, ou zéro.</returns>
[DllImport("user32.dll", EntryPoint = "GetParent")]
        private static extern IntPtr NativeObjectBrowserParent(IntPtr handle);

        /// <summary>Lit l’état d’activation natif d’une fenêtre.</summary>
        /// <param name="handle">Handle de fenêtre à tester.</param>
        /// <returns><see langword="true"/> si Windows indique que la fenêtre est activée.</returns>
[DllImport("user32.dll", EntryPoint = "IsWindowEnabled")]
        private static extern bool NativeObjectBrowserEnabled(IntPtr handle);

        /// <summary>Native ancestry and enabled-state boundaries shared by the browser and code view.</summary>
        internal static Func<IntPtr, IntPtr> ObjectBrowserParent = NativeObjectBrowserParent;
        /// <summary>Vérificateur de disponibilité native des ancêtres des contrôles.</summary>
internal static Func<IntPtr, bool> ObjectBrowserEnabled = NativeObjectBrowserEnabled;

        /// <summary>Sélectionne une bibliothèque, une classe ou un membre dans le navigateur déjà ouvert.</summary>
        /// <param name="request">Bibliothèque facultative, nom de classe et nom de membre éventuel.</param>
        /// <returns>État de sélection relu et instantané du navigateur.</returns>
internal static object SelectObjectBrowser(Request request)
        {
            if ((request.ObjectName == null && request.Context == null) ||
                (request.ObjectName != null && (string.IsNullOrWhiteSpace(request.ObjectName) || request.ObjectName.Length > 255)) ||
                (request.Procedure != null && (request.ObjectName == null || string.IsNullOrWhiteSpace(request.Procedure) || request.Procedure.Length > 255)) ||
                (request.Context != null && (string.IsNullOrWhiteSpace(request.Context) || request.Context.Length > 255)))
                throw new ArgumentException("ObjectName or library Context is required; Procedure requires ObjectName. Values must contain 1-255 characters.");
            IntPtr root = FindVbeRoot();
            IntPtr pane = root == IntPtr.Zero ? IntPtr.Zero : FindPane(ChildWindows(root), "Object Browser", "Explorateur d'objets");
            if (pane == IntPtr.Zero) throw new InvalidOperationException("Open the Object Browser before selecting a class.");
            for (IntPtr ancestor = pane; ancestor != IntPtr.Zero; ancestor = ObjectBrowserParent(ancestor))
                if (!ObjectBrowserEnabled(ancestor))
                    throw new InvalidOperationException("Object Browser navigation is disabled by a modal window. Read debug_dialog before selecting an item.");
            var controls = ObjectBrowserControls(pane);
            if (request.Context != null)
            {
                SelectBrowserLibrary(controls, request.Context);
                controls = ObjectBrowserControls(pane);
            }
            if (request.ObjectName != null) SelectBrowserLabel(controls, new[] { "Classes " + request.ObjectName });
            // LBN_SELCHANGE is asynchronous. Wait for the member list belonging to
            // the requested class; never select a same-named member of the old class.
            if (!string.IsNullOrEmpty(request.Procedure))
            {
                string[] labels = { "Membres de '" + request.ObjectName + "' " + request.Procedure,
                    "Members of '" + request.ObjectName + "' " + request.Procedure };
                bool found = false;
                for (int attempt = 0; attempt < 10; attempt++)
                {
                    PauseNative(100);
                    if (SelectBrowserLabel(controls, labels, true)) { found = true; break; }
                }
                if (!found) throw new InvalidOperationException("Class selection was sent, but the requested member was not found. Read the current browser before retrying.");
            }
            PauseNative(100);
            var selectedLabels = new List<string>();
            foreach (AutomationElement control in controls)
                if (control.Current.ControlType == ControlType.List && control.TryGetCurrentPattern(SelectionPattern.Pattern, out object current))
                    selectedLabels.AddRange(((SelectionPattern)current).Current.GetSelection().Select(item => item.Current.Name));
            bool libraryObserved = request.Context == null || controls.Any(control =>
                control.Current.ControlType == ControlType.ComboBox &&
                (control.Current.Name == "Bibliothèques" || control.Current.Name == "Libraries") &&
                control.TryGetCurrentPattern(SelectionPattern.Pattern, out object librarySelection) &&
                ((SelectionPattern)librarySelection).Current.GetSelection().Any(item => item.Current.Name == request.Context));
            bool observed = libraryObserved && (request.ObjectName == null || selectedLabels.Contains("Classes " + request.ObjectName)) &&
                (string.IsNullOrEmpty(request.Procedure) ||
                 selectedLabels.Contains("Membres de '" + request.ObjectName + "' " + request.Procedure) ||
                 selectedLabels.Contains("Members of '" + request.ObjectName + "' " + request.Procedure));
            return new { SelectionObserved = observed, ObjectName = request.ObjectName, Procedure = request.Procedure, Library = request.Context,
                Snapshot = ReadObjectBrowser(), SemanticResolutionVerified = false };
        }

        /// <summary>Liste par pages les bibliothèques, classes ou membres exposés par le navigateur natif.</summary>
        /// <param name="request">Type de volet, filtre éventuel, décalage et taille de page.</param>
        /// <returns>Éléments accessibles et informations de pagination.</returns>
internal static object ListObjectBrowser(Request request)
        {
            if (request.Pane != "classes" && request.Pane != "members" && request.Pane != "libraries")
                throw new ArgumentException("Pane must be classes, members or libraries.");
            if (request.Offset < 0 || request.Limit < 0 || request.Limit > 200)
                throw new ArgumentException("Offset must be nonnegative; Limit must be 0 (default 50) or 1-200.");
            IntPtr root = FindVbeRoot();
            IntPtr pane = root == IntPtr.Zero ? IntPtr.Zero : FindPane(ChildWindows(root), "Object Browser", "Explorateur d'objets");
            if (pane == IntPtr.Zero) throw new InvalidOperationException("Open the Object Browser before listing its items.");
            var candidates = new List<AutomationElement>();
            foreach (var control in ObjectBrowserControls(pane))
            {
                if (request.Pane == "libraries")
                {
                    if (control.Current.ControlType == ControlType.ComboBox &&
                        (control.Current.Name == "Bibliothèques" || control.Current.Name == "Libraries")) candidates.Add(control);
                }
                else if (control.Current.ControlType == ControlType.List)
                {
                    var first = control.FindFirst(TreeScope.Children, new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.ListItem));
                    string label = first == null ? "" : first.Current.Name;
                    if (request.Pane == "classes" && label.StartsWith("Classes ", StringComparison.Ordinal)) candidates.Add(control);
                    if (request.Pane == "members" && (label.StartsWith("Membres de '", StringComparison.Ordinal) || label.StartsWith("Members of '", StringComparison.Ordinal))) candidates.Add(control);
                }
            }
            if (candidates.Count != 1) return new { Available = false, Reason = "The requested nonempty native list was not uniquely identified." };
            var list = candidates[0];
            var items = list.FindAll(request.Pane == "libraries" ? TreeScope.Descendants : TreeScope.Children,
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.ListItem));
            var rows = new List<object>();
            int matched = 0;
            int limit = request.Limit == 0 ? 50 : request.Limit;
            foreach (AutomationElement item in items)
            {
                string label = item.Current.Name;
                if (!string.IsNullOrEmpty(request.Query) && label.IndexOf(request.Query, StringComparison.OrdinalIgnoreCase) < 0) continue;
                if (matched++ < request.Offset || rows.Count >= limit) continue;
                bool readable = item.TryGetCurrentPattern(SelectionItemPattern.Pattern, out object selection);
                rows.Add(new { Label = label, SelectionReadable = readable,
                    Selected = readable && ((SelectionItemPattern)selection).Current.IsSelected });
            }
            return new { Available = true, Pane = request.Pane, Query = request.Query, Offset = request.Offset,
                Limit = limit, TotalMatches = matched, HasMore = matched - request.Offset > rows.Count,
                Items = rows, Source = "Native VBE UI Automation", Enabled = list.Current.IsEnabled };
        }

        /// <summary>Sélectionne une bibliothèque exacte puis vérifie l’état de la combo native.</summary>
        /// <param name="controls">Contrôles UI Automation du navigateur courant.</param>
        /// <param name="library">Nom exact de la bibliothèque à sélectionner.</param>
private static void SelectBrowserLibrary(IEnumerable<AutomationElement> controls, string library)
        {
            var combos = controls.Where(control => control.Current.ControlType == ControlType.ComboBox &&
                (control.Current.Name == "Bibliothèques" || control.Current.Name == "Libraries")).ToArray();
            if (combos.Length != 1) throw new InvalidOperationException("The library selector was not uniquely identified.");
            var combo = combos[0];
            IntPtr handle = new IntPtr(combo.Current.NativeWindowHandle);
            if (handle == IntPtr.Zero || ClassName(handle) != "ComboBox")
                throw new InvalidOperationException("The native library combo was not identified.");
            for (IntPtr ancestor = handle; ancestor != IntPtr.Zero; ancestor = ObjectBrowserParent(ancestor))
                if (!ObjectBrowserEnabled(ancestor))
                    throw new InvalidOperationException("The native library selector is disabled by ancestor " + ClassName(ancestor) + " (" + WindowText(ancestor) + ").");
            var items = combo.FindAll(TreeScope.Descendants, new AndCondition(
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.ListItem),
                new PropertyCondition(AutomationElement.NameProperty, library)));
            if (items.Count != 1) throw new InvalidOperationException("Expected one exact library; found " + items.Count + ".");
            if (!items[0].TryGetCurrentPattern(SelectionItemPattern.Pattern, out object selection))
                throw new InvalidOperationException("The library item cannot be selected.");
            // Use the native accessibility action so VBE receives the complete
            // combo selection lifecycle, not only a changed displayed index.
            ((SelectionItemPattern)selection).Select();
            // handle was validated above and cannot change during the accessibility action.
            if (!PostMessage(ObjectBrowserParent(handle), 0x111,
                new IntPtr(GetDlgCtrlID(handle) | (1 << 16)), handle))
                throw new InvalidOperationException("Library selection changed but notification failed. Read the browser before retrying.");
            PauseNative(150);
            if (!combo.TryGetCurrentPattern(SelectionPattern.Pattern, out object current) ||
                !((SelectionPattern)current).Current.GetSelection().Any(item => item.Current.Name == library))
                throw new InvalidOperationException("Native library selection could not be verified.");
        }

        /// <summary>Sélectionne un élément de liste par une ou plusieurs légendes exactes.</summary>
        /// <param name="controls">Contrôles de liste du navigateur.</param>
        /// <param name="labels">Légendes de langue ou de classe acceptées.</param>
        /// <param name="allowMissing">Autorise le retour faux lorsqu’aucun élément n’est encore apparu.</param>
        /// <returns><see langword="true"/> si l’élément unique a été sélectionné et notifié.</returns>
private static bool SelectBrowserLabel(IEnumerable<AutomationElement> controls, string[] labels, bool allowMissing = false)
        {
            var matches = new List<Tuple<AutomationElement, AutomationElement>>();
            foreach (AutomationElement control in controls)
            {
                if (control.Current.ControlType != ControlType.List) continue;
                var condition = labels.Length == 1
                    ? (Condition)new PropertyCondition(AutomationElement.NameProperty, labels[0])
                    : new OrCondition(labels.Select(label => (Condition)new PropertyCondition(AutomationElement.NameProperty, label)).ToArray());
                foreach (AutomationElement item in control.FindAll(TreeScope.Children, condition))
                    matches.Add(Tuple.Create(control, item));
            }
            if (matches.Count == 0 && allowMissing) return false;
            if (matches.Count != 1) throw new InvalidOperationException("Expected one exact Object Browser item; found " + matches.Count + ". Supply a narrower library Context if ambiguous.");
            var match = matches[0];
            IntPtr handle = new IntPtr(match.Item1.Current.NativeWindowHandle);
            IntPtr parent = ObjectBrowserParent(handle);
            if (handle == IntPtr.Zero || parent == IntPtr.Zero || !match.Item2.TryGetCurrentPattern(SelectionItemPattern.Pattern, out object pattern))
                throw new InvalidOperationException("The native browser item cannot be selected.");
            var selection = (SelectionItemPattern)pattern;
            selection.Select();
            if (!selection.Current.IsSelected) throw new InvalidOperationException("Native selection readback failed.");
            int notification = GetDlgCtrlID(handle) | (1 << 16);
            if (!PostMessage(parent, 0x111, new IntPtr(notification), handle))
                throw new InvalidOperationException("Selection changed, but its native notification failed. Read the browser before retrying.");
            return true;
        }
    }
}
