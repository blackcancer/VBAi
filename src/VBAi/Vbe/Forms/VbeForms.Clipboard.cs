using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;

namespace VBAi
{

    /// <summary>Expose le presse-papiers natif du Designer avec versions de précondition et récupération de coupes.</summary>
    internal sealed partial class VbeForms
    {

        /// <summary>Données de récupération d’une coupe Designer encore disponible dans la session.</summary>
        private sealed class ClipboardRecovery
        {

            /// <summary>Identifiant de récupération et chemin canonique du conteneur auxquels appartiennent les données copiées.</summary>
            public string Id, ParentPath;

            /// <summary>Référence du UserForm vivant auquel cette récupération est attachée.</summary>
            public object Form;

            /// <summary>Copie bornée des formats MSForms capturés avant la coupe.</summary>
            public DesignerClipboardBackup Backup;

            /// <summary>Arbre avant la coupe et arbre après une coupe ayant produit des différences observables.</summary>
            public object OriginalTree, CutTree;

            /// <summary>Géométrie et noms des contrôles sélectionnés lors de la coupe, conservés pour le rapport de récupération.</summary>
            public FormLayoutBox[] Boxes;

            /// <summary>Ordre des indices TabIndex et noms des contrôles présents au moment de la coupe.</summary>
            public string[] TabOrder;

            /// <summary>Prevents retrying recovery after a native clipboard write with an uncertain outcome.</summary>
            public bool RecoveryAttempted;
        }

        /// <summary>Conserve les huit sauvegardes de coupe les plus récentes de la session.</summary>
        private readonly Queue<ClipboardRecovery> clipboardRecoveries = new Queue<ClipboardRecovery>();

        /// <summary>Republie les formats sauvegardés après vérification des versions du Designer et du presse-papiers.</summary>
        /// <param name="request">Formulaire, conteneur et révisions attendues avec identifiant de récupération.</param>
        /// <returns>État de relecture des données republiées et prochaine action explicite proposée.</returns>
        /// <exception cref="InvalidOperationException">La sauvegarde a expiré ou ne correspond pas au formulaire courant.</exception>
        public object RestoreDesignerClipboard(Request request)
        {
            dynamic form = GetForm(GetDesignProject(request.Project), request.Form);
            var recovery = clipboardRecoveries.FirstOrDefault(x => x.Id == request.DesignerClipboardRecoveryId) ?? throw new InvalidOperationException("Clipboard recovery not found or expired; only the latest 8 cuts in this session are retained.");
            if (!ReferenceEquals((object)form, recovery.Form) || (request.ParentPath ?? "") != recovery.ParentPath)
                throw new InvalidOperationException("Clipboard recovery belongs to another live form or container.");
            dynamic before = ClipboardState(request.Project, request.Form, request.ParentPath);
            RequireClipboardRevision(request, (string)before.SelectionVersion, (string)before.ClipboardVersion);
            WriteDesignerClipboard(recovery.Backup.CreateDataObject(), true);
            bool verified = recovery.Backup.Matches(ReadDesignerClipboard());
            return new
            {
                Restored = verified,
                recovery.Backup.OmittedFormats,
                State = ClipboardState(request.Project, request.Form, request.ParentPath),
                NextAction = "native_form_clipboard paste",
                Limit = "Restores captured clipboard formats only; does not recreate controls. Inspect the current tree before an explicit paste to avoid duplicates. Native paste may change placement; omitted formats are not restored."
            };
        }

        /// <summary>Lit le numéro de séquence natif du presse-papiers Windows.</summary>
        /// <returns>Numéro incrémenté lors des changements du presse-papiers.</returns>
        [DllImport("user32.dll", EntryPoint = "GetClipboardSequenceNumber")] private static extern uint NativeClipboardSequence();

        /// <summary>Native clipboard boundaries; tests retain ownership of an in-memory IDataObject only.</summary>
        internal static Func<uint> DesignerClipboardSequence = NativeClipboardSequence;

        /// <summary>Délégué injectable de lecture du presse-papiers MSForms.</summary>
        internal static Func<System.Windows.Forms.IDataObject> ReadDesignerClipboard = System.Windows.Forms.Clipboard.GetDataObject;

        /// <summary>Délégué injectable d’écriture du presse-papiers MSForms.</summary>
        internal static Action<System.Windows.Forms.IDataObject, bool> WriteDesignerClipboard = System.Windows.Forms.Clipboard.SetDataObject;

        /// <summary>Capture l’arbre, la sélection courante et le numéro de séquence du presse-papiers sans lire ses données binaires.</summary>
        /// <param name="projectName">Projet VBA contenant le formulaire.</param>
        /// <param name="formName">Nom du UserForm inspecté.</param>
        /// <param name="parentPath">Chemin canonique d’un conteneur direct, ou null pour le formulaire racine.</param>
        /// <returns>Arbre, noms sélectionnés, version de sélection, version de presse-papiers et capacité de collage.</returns>
        /// <exception cref="InvalidOperationException">Le presse-papiers a changé pendant l’inspection ou la sélection est trop grande.</exception>
        public object ClipboardState(string projectName, string formName, string parentPath = null)
        {
            dynamic form = GetForm(GetDesignProject(projectName), formName);
            dynamic tree = Tree(projectName, formName);
            dynamic container = ClipboardContainer((object)form.Designer, (object)tree, parentPath);
            var selected = new List<string>();
            foreach (dynamic control in container.Selected)
            {
                if (selected.Count >= 256) throw new InvalidOperationException("Designer selection exceeds 256 controls.");
                selected.Add((string)control.Name);
            }
            uint sequence = DesignerClipboardSequence();
            bool canPaste = (bool)container.CanPaste;
            if (sequence != DesignerClipboardSequence()) throw new InvalidOperationException("Clipboard changed during inspection.");
            string selectionVersion = VbeCodeClipboard.Hash(new System.Web.Script.Serialization.JavaScriptSerializer().Serialize(
                new { Project = projectName, Form = formName, ParentPath = parentPath ?? "", TreeVersion = (string)tree.TreeVersion, Selected = selected }));
            return new
            {
                Project = projectName,
                Form = formName,
                ParentPath = parentPath ?? "",
                Tree = (object)tree,
                Selected = selected,
                SelectionVersion = selectionVersion,
                ClipboardVersion = sequence.ToString(CultureInfo.InvariantCulture),
                CanPaste = canPaste,
                Limit = "Current Designer selection; binary clipboard data is not read or sent to the model. Clipboard revision is a Windows sequence number, not a content hash."
            };
        }

        /// <summary>Exécute Copy, Cut ou Paste sur le conteneur courant avec validation des versions Designer et clipboard.</summary>
        /// <param name="request">Projet, formulaire, action, conteneur et révisions lues précédemment.</param>
        /// <returns>Différences observées, évolution du numéro clipboard et récupération éventuelle de la coupe.</returns>
        /// <exception cref="ArgumentException">L’action demandée n’est pas copy, cut ou paste.</exception>
        /// <exception cref="InvalidOperationException">Aucune sélection n’existe ou le conteneur ne peut pas coller.</exception>
        public object NativeClipboard(Request request)
        {
            if (request.Action != "copy" && request.Action != "cut" && request.Action != "paste")
                throw new ArgumentException("Action must be copy, cut or paste.");
            dynamic form = GetForm(GetDesignProject(request.Project), request.Form);
            dynamic before = ClipboardState(request.Project, request.Form, request.ParentPath);
            RequireClipboardRevision(request, (string)before.SelectionVersion, (string)before.ClipboardVersion);
            if (request.Action != "paste" && ((List<string>)before.Selected).Count == 0)
                throw new InvalidOperationException("Select Designer controls before copying or cutting.");
            if (request.Action == "paste" && !(bool)before.CanPaste)
                throw new InvalidOperationException("The Designer cannot paste the current clipboard.");
            dynamic container = ClipboardContainer((object)form.Designer, (object)before.Tree, request.ParentPath);
            string error = null; object after = null; ClipboardRecovery recovery = null;
            try
            {
                if (request.Action == "copy") container.Copy();
                else if (request.Action == "cut")
                {
                    // Capture an eager copy before removal; no delayed COM clipboard object is retained.
                    var boxes = new List<FormLayoutBox>(); var tabs = new SortedDictionary<int, string>();
                    foreach (dynamic control in container.Controls)
                    {
                        string name = (string)control.Name; tabs.Add((int)control.TabIndex, name);
                        if (((List<string>)before.Selected).Contains(name)) boxes.Add(new FormLayoutBox
                        {
                            Path = name,
                            Left = (double)control.Left,
                            Top = (double)control.Top,
                            Width = (double)control.Width,
                            Height = (double)control.Height
                        });
                    }
                    container.Copy();
                    uint sequence = DesignerClipboardSequence();
                    var backup = DesignerClipboardBackup.Capture(ReadDesignerClipboard());
                    if (sequence != DesignerClipboardSequence()) throw new InvalidOperationException("Clipboard changed while preparing recovery; cut was not attempted.");
                    recovery = new ClipboardRecovery { Id = Guid.NewGuid().ToString("N"), Form = (object)form, ParentPath = request.ParentPath ?? "", Backup = backup, OriginalTree = (object)before.Tree, Boxes = boxes.ToArray(), TabOrder = tabs.Values.ToArray() };
                    if (clipboardRecoveries.Count == 8) clipboardRecoveries.Dequeue();
                    clipboardRecoveries.Enqueue(recovery);
                    container.Cut();
                }
                else container.Paste();
            }
            catch (Exception ex) { error = ex.Message; }
            try { after = ClipboardState(request.Project, request.Form, request.ParentPath); }
            catch (Exception ex) { error = error ?? ex.Message; }
            object afterTree = after == null ? null : (object)((dynamic)after).Tree;
            var changes = afterTree == null ? new FormHistoryDiff.Change[0] : FormHistoryDiff.Compare((object)before.Tree, afterTree);
            if (recovery != null && afterTree != null && error == null && changes.Length > 0) recovery.CutTree = afterTree;
            bool clipboardChanged = after != null && (string)((dynamic)after).ClipboardVersion != (string)before.ClipboardVersion;
            return new
            {
                request.Project,
                request.Form,
                request.Action,
                Executed = error == null,
                DesignerChangeObserved = changes.Length > 0,
                ClipboardChanged = clipboardChanged,
                DesignerClipboardRecoveryId = recovery?.Id,
                RecoveryOmittedFormats = recovery?.Backup.OmittedFormats,
                RecoveryBytes = recovery?.Backup.ByteCount,
                ClipboardContentVerified = false,
                Before = (object)before,
                After = after,
                DesignerChanges = changes,
                ReadErrorsBefore = FormHistoryDiff.ReadErrorCount((object)before.Tree),
                ReadErrorsAfter = afterTree == null ? (int?)null : FormHistoryDiff.ReadErrorCount(afterTree),
                NativeError = error,
                Saved = false,
                NextRead = "form_clipboard_state",
                Limit = "Native selected controls are copied/cut; paste uses this Designer's current selection context. Event-handler code is not transferred. Binary clipboard contents are not inspected. Partial changes are reported without retry or implicit rollback; Copy/Cut/Paste may not enter native undo history; inspect CanUndo before offering native_form_history. Cut stores up to 8 MiB of readable MSForms formats before removal (latest 8 cuts in this session). restore_form_clipboard can republish the recovery; it does not undo or restore geometry automatically."
            };
        }

        /// <summary>Résout le formulaire racine ou un conteneur canonique pouvant exposer Controls.</summary>
        /// <param name="designer">Designer du UserForm.</param>
        /// <param name="tree">Arbre de formulaire déjà lu.</param>
        /// <param name="parentPath">Chemin du conteneur direct, vide ou nul pour la racine.</param>
        /// <returns>Objet COM du conteneur validé.</returns>
        /// <exception cref="ArgumentException">Le chemin du parent n’est pas canonique.</exception>
        private static object ClipboardContainer(object designer, object tree, string parentPath)
        {
            if (string.IsNullOrEmpty(parentPath)) return designer;
            if (!TreeContainsPath((IEnumerable)((dynamic)tree).Controls, parentPath))
                throw new ArgumentException("ParentPath is not canonical in form_tree.");
            // Validate it exposes a child Controls collection, excluding leaf controls.
            ResolveNestedControls((dynamic)designer, parentPath);
            return ResolveTreeItem(designer, parentPath);
        }

        /// <summary>Remplace la sélection dans un conteneur par les noms directs fournis, sans changer le focus ni le presse-papiers.</summary>
        /// <param name="request">Conteneur, noms sélectionnés et version de sélection attendue.</param>
        /// <returns>État avant/après et indication de vérification de la sélection.</returns>
        /// <exception cref="ArgumentException">Les noms ne sont pas uniques ou n’appartiennent pas tous au conteneur.</exception>
        /// <exception cref="InvalidOperationException">La sélection ou l’arbre a changé depuis sa lecture.</exception>
        public object SelectDesignerControls(Request request)
        {
            if (request.Items == null || request.Items.Length > 64 ||
                request.Items.Any(string.IsNullOrWhiteSpace) || request.Items.Distinct(StringComparer.Ordinal).Count() != request.Items.Length)
                throw new ArgumentException("Items must contain 0 to 64 unique direct control names; empty clears this container's selection.");
            dynamic form = GetForm(GetDesignProject(request.Project), request.Form);
            dynamic before = ClipboardState(request.Project, request.Form, request.ParentPath);
            if (string.IsNullOrWhiteSpace(request.ExpectedDesignerSelectionVersion) ||
                !string.Equals(request.ExpectedDesignerSelectionVersion, (string)before.SelectionVersion, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("Designer tree or selection changed; read form_clipboard_state again.");
            dynamic container = ClipboardContainer((object)form.Designer, (object)before.Tree, request.ParentPath);
            var controls = new Dictionary<string, object>(StringComparer.Ordinal);
            foreach (dynamic control in container.Controls) controls.Add((string)control.Name, (object)control);
            if (request.Items.Any(name => !controls.ContainsKey(name)))
                throw new ArgumentException("Items contains a name outside the selected container; no selection changed.");
            string error = null; object after = null;
            try
            {
                foreach (var entry in controls) ((dynamic)entry.Value).InSelection = request.Items.Contains(entry.Key, StringComparer.Ordinal);
            }
            catch (Exception ex) { error = ex.Message; }
            try { after = ClipboardState(request.Project, request.Form, request.ParentPath); }
            catch (Exception ex) { error = error ?? ex.Message; }
            bool verified = after != null && error == null &&
                new HashSet<string>((List<string>)((dynamic)after).Selected, StringComparer.Ordinal).SetEquals(request.Items);
            return new
            {
                Verified = verified,
                Before = (object)before,
                After = after,
                NativeError = error,
                NextRead = "form_clipboard_state",
                Limit = "Changes only the specified container selection; selections in other containers are independent. No focus or clipboard mutation. On partial failure re-read before another action."
            };
        }

        /// <summary>Refuse une commande lorsque la sélection Designer ou le presse-papiers diffère de la lecture attendue.</summary>
        /// <param name="request">Préconditions transmises avec la commande.</param>
        /// <param name="selectionVersion">Version courante de l’arbre et de la sélection.</param>
        /// <param name="clipboardVersion">Numéro courant de séquence Windows.</param>
        /// <exception cref="InvalidOperationException">Une version attendue est absente, nulle ou ne correspond plus.</exception>
        internal static void RequireClipboardRevision(Request request, string selectionVersion, string clipboardVersion)
        {
            if (string.IsNullOrWhiteSpace(request.ExpectedDesignerSelectionVersion) ||
                !string.Equals(request.ExpectedDesignerSelectionVersion, selectionVersion, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("Designer tree or selection changed; read form_clipboard_state again.");
            if (string.IsNullOrWhiteSpace(request.ExpectedClipboardVersion) || request.ExpectedClipboardVersion != clipboardVersion || clipboardVersion == "0")
                throw new InvalidOperationException("Clipboard changed or is unavailable; read form_clipboard_state again.");
        }
    }
}
