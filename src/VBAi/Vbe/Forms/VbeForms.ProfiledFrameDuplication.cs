using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.Linq;

namespace VBAi
{

    /// <summary>Implémente des profils positifs et bornés pour copier une Frame racine avec certains enfants directs.</summary>
    internal sealed partial class VbeForms
    {

        /// <summary>Liste les champs et limites vérifiés pour un type de contrôle MSForms.</summary>
        private sealed class VerifiedChildProfile
        {

            /// <summary>Nom du type MSForms visé par le profil.</summary>
            public string Type;

            /// <summary>ProgID utilisé pour créer le contrôle.</summary>
            public string ProgId;

            /// <summary>Noms des propriétés que le profil copie et vérifie.</summary>
            public string[] Fields;

            /// <summary>Propriétés et comportements exclus de ce profil.</summary>
            public string Limitation;
        }

        /// <summary>Capture le profil et les valeurs lues d’un enfant avant sa duplication.</summary>
        private sealed class ProfiledChildSnapshot
        {

            /// <summary>Profil positif utilisé pour lire, écrire et vérifier l’enfant.</summary>
            public VerifiedChildProfile Profile;

            /// <summary>Chemin canonique du contrôle source.</summary>
            public string SourcePath;

            /// <summary>Nom du contrôle à créer.</summary>
            public string ProposedName;

            /// <summary>Chemin qui identifiera le contrôle créé.</summary>
            public string ProposedPath;

            /// <summary>Légende d’un Label, CheckBox, CommandButton ou OptionButton.</summary>
            public string Caption;

            /// <summary>Valeur textuelle du TextBox, éventuellement absente.</summary>
            public string TextValue;

            /// <summary>État booléen du CheckBox source.</summary>
            public bool BooleanValue;

            /// <summary>Largeur de liste textuelle du ComboBox.</summary>
            public string ListWidth;

            /// <summary>Position horizontale de l’enfant.</summary>
            public double Left;

            /// <summary>Position verticale de l’enfant.</summary>
            public double Top;

            /// <summary>Largeur de l’enfant.</summary>
            public double Width;

            /// <summary>Hauteur de l’enfant.</summary>
            public double Height;

            /// <summary>Couleur OLE du fond d’un Label.</summary>
            public int BackColor;

            /// <summary>Nom de police d’un Label.</summary>
            public string FontName;

            /// <summary>Taille de police d’un Label.</summary>
            public double FontSize;

            /// <summary>Indique si la police d’un Label est en gras.</summary>
            public bool FontBold;
        }

        /// <summary>Capture la Frame, sa version d’arbre, ses enfants admissibles et les problèmes de prévalidation.</summary>
        private sealed class ProfiledFrameSnapshot
        {

            /// <summary>Chemin de la Frame source.</summary>
            public string SourcePath;

            /// <summary>Nom proposé pour la nouvelle Frame.</summary>
            public string NewName;

            /// <summary>Légende de la Frame.</summary>
            public string Caption;

            /// <summary>Position horizontale de la Frame.</summary>
            public double Left;

            /// <summary>Position verticale de la Frame.</summary>
            public double Top;

            /// <summary>Largeur de la Frame.</summary>
            public double Width;

            /// <summary>Hauteur de la Frame.</summary>
            public double Height;

            /// <summary>Version d’arbre utilisée pour contrôler la prévalidation.</summary>
            public string TreeVersion;

            /// <summary>Nombre de nœuds du formulaire avant copie.</summary>
            public int NodeCount;

            /// <summary>Snapshots des enfants directs qui disposent d’un profil lisible.</summary>
            public readonly List<ProfiledChildSnapshot> Children = new List<ProfiledChildSnapshot>();

            /// <summary>Motifs qui rendent le profil proposé inadmissible.</summary>
            public readonly List<string> Issues = new List<string>();
        }

        // A positive registry. No profile is inferred from COM IsReadOnly or
        // PROPERTYPUT metadata; every field below came from a prior Excel probe.
        /// <summary>Profils positifs limités aux types contrôlés : Label, TextBox, CheckBox, CommandButton, ComboBox et OptionButton.</summary>
        private static readonly Dictionary<string, VerifiedChildProfile> FrameChildProfiles =
            new Dictionary<string, VerifiedChildProfile>(StringComparer.OrdinalIgnoreCase)
            {
                ["Label"] = new VerifiedChildProfile { Type = "Label", ProgId = "Forms.Label.1",
                    Fields = new[] { "Name", "Caption", "Left", "Top", "Width", "Height",
                        "BackColor", "Font.Name", "Font.Size", "Font.Bold" },
                    Limitation = "Other Label properties, image, and z-order are not copied." },
                ["TextBox"] = new VerifiedChildProfile { Type = "TextBox", ProgId = "Forms.TextBox.1",
                    Fields = new[] { "Name", "Left", "Top", "Width", "Height", "Value (text or empty)" },
                    Limitation = "Bindings, validation, and other TextBox properties are not copied." },
                ["CheckBox"] = new VerifiedChildProfile { Type = "CheckBox", ProgId = "Forms.CheckBox.1",
                    Fields = new[] { "Name", "Caption", "Left", "Top", "Width", "Height",
                        "Value (Boolean only)" },
                    Limitation = "TriState/null and other CheckBox properties are not copied." },
                ["CommandButton"] = new VerifiedChildProfile { Type = "CommandButton", ProgId = "Forms.CommandButton.1",
                    Fields = new[] { "Name", "Caption", "Left", "Top", "Width", "Height" },
                    Limitation = "Event procedures and other CommandButton properties are not copied." },
                ["ComboBox"] = new VerifiedChildProfile { Type = "ComboBox", ProgId = "Forms.ComboBox.1",
                    Fields = new[] { "Name", "Left", "Top", "Width", "Height", "ListWidth (text)" },
                    Limitation = "Items, bindings, selection, and other ComboBox properties are not copied." },
                ["OptionButton"] = new VerifiedChildProfile { Type = "OptionButton", ProgId = "Forms.OptionButton.1",
                    Fields = new[] { "Name", "Caption", "Left", "Top", "Width", "Height" },
                    Limitation = "Value and GroupName are not copied; copying selection can affect siblings." }
            };

        /// <summary>Retourne en lecture seule les profils d’enfants reconnus et les motifs d’inadmissibilité.</summary>
        /// <param name="request">Projet, formulaire, chemin de Frame, version attendue et nom proposé.</param>
        /// <returns>Plan sérialisable qui indique les champs copiés, les limites et l’absence de mutation.</returns>
        public object FrameProfileCopyPlan(Request request)
        {
            ProfiledFrameSnapshot plan = ReadFrameProfilePlan(request);
            return new { Project = request.Project, Form = request.Form,
                SourcePath = plan.SourcePath, ProposedFrameName = plan.NewName,
                ExpectedTreeVersion = plan.TreeVersion, DirectChildCount = plan.Children.Count,
                Children = plan.Children.Select(child => new { SourcePath = child.SourcePath,
                    ProposedPath = child.ProposedPath, Type = child.Profile.Type,
                    CopiedFields = child.Profile.Fields, Limitation = child.Profile.Limitation }).ToArray(),
                Issues = plan.Issues, EligibleForLimitedProbe = plan.Issues.Count == 0,
                Completeness = "Partial", MutationVerified = false, ReadOnly = true,
                Scope = "Root Frame and direct leaf children with explicit positive profiles only." };
        }

        /// <summary>Valide la Frame racine et sa version puis capture les enfants directs admissibles sans muter le formulaire.</summary>
        /// <param name="request">Données du projet, du formulaire, de la Frame et du nom proposé.</param>
        /// <returns>Snapshot de prévalidation avec les enfants et problèmes observés.</returns>
        /// <exception cref="ArgumentException">Le chemin ne vise pas une Frame racine ou la version attendue manque.</exception>
        /// <exception cref="InvalidOperationException">La hiérarchie a changé ou le contrôle ne correspond pas à une Frame native.</exception>
        private ProfiledFrameSnapshot ReadFrameProfilePlan(Request request)
        {
            if (string.IsNullOrWhiteSpace(request.ControlPath) ||
                request.ControlPath.Split('/').Length != 2 ||
                !request.ControlPath.StartsWith("Controls/", StringComparison.Ordinal))
                throw new ArgumentException("Only a root Controls/<Frame> path is supported.");
            if (string.IsNullOrWhiteSpace(request.ExpectedTreeVersion))
                throw new ArgumentException("ExpectedTreeVersion is required from form_tree.");
            ValidateName(request.NewName, "NewName");
            dynamic form = GetForm(GetProject(request.Project), request.Form);
            dynamic tree = Tree(request.Project, request.Form);
            if (!string.Equals((string)tree.TreeVersion, request.ExpectedTreeVersion,
                StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The UserForm hierarchy changed since it was read.");
            if (!TreeContainsPath((IEnumerable)tree.Controls, request.ControlPath))
                throw new InvalidOperationException("ControlPath is not canonical in form_tree.");
            object target = ResolveTreeItem(form.Designer, request.ControlPath);
            if (!string.Equals(TypeDescriptor.GetClassName(target), "Frame", StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The selected control is not a native MSForms Frame.");

            dynamic frame = target;
            var plan = new ProfiledFrameSnapshot {
                SourcePath = request.ControlPath, NewName = request.NewName,
                TreeVersion = (string)tree.TreeVersion, NodeCount = (int)tree.NodeCount,
                Caption = (string)frame.Caption,
                Left = Convert.ToDouble(frame.Left, CultureInfo.InvariantCulture),
                Top = Convert.ToDouble(frame.Top, CultureInfo.InvariantCulture),
                Width = Convert.ToDouble(frame.Width, CultureInfo.InvariantCulture),
                Height = Convert.ToDouble(frame.Height, CultureInfo.InvariantCulture)
            };
            try { ValidateCopyBox(plan.Left, plan.Top, plan.Width, plan.Height, "Frame"); }
            catch (Exception ex) { plan.Issues.Add(ex.Message); }
            if (request.NewName.Length > 40)
                plan.Issues.Add("New Frame name exceeds 40 characters.");

            var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (dynamic existing in form.Designer.Controls)
                names.Add((string)existing.Name);
            if (names.Contains(request.NewName))
                plan.Issues.Add("New Frame name already exists: " + request.NewName);
            int directCount = 0;
            foreach (dynamic child in frame.Controls)
            {
                if (!SameContainer((object)child.Parent, target, (string)form.Name)) continue;
                directCount++;
                if (directCount > 128)
                    throw new InvalidOperationException("Frame exceeds the 128 direct-child copy limit.");
                string sourceName = (string)child.Name;
                string type = TypeDescriptor.GetClassName((object)child);
                VerifiedChildProfile profile;
                if (!FrameChildProfiles.TryGetValue(type, out profile))
                {
                    plan.Issues.Add(sourceName + ": no positive copy profile for " + type + ".");
                    continue;
                }
                string newName = request.NewName + "_" + sourceName;
                if (newName.Length > 40)
                    plan.Issues.Add(sourceName + ": proposed child name exceeds 40 characters.");
                if (!names.Add(newName))
                    plan.Issues.Add(sourceName + ": proposed child name already exists: " + newName);
                try
                {
                    plan.Children.Add(ReadProfiledChild(child, profile,
                        request.ControlPath + "/Controls/" + sourceName,
                        "Controls/" + request.NewName + "/Controls/" + newName, newName));
                }
                catch (Exception ex)
                {
                    plan.Issues.Add(sourceName + ": " + ex.Message);
                }
            }
            if (directCount == 0)
                plan.Issues.Add("Use duplicate_empty_form_frame for a Frame without children.");
            if (plan.Children.Count != directCount)
                plan.Issues.Add("Not all direct children have readable positive profiles.");
            return plan;
        }

        /// <summary>Lit les valeurs autorisées du contrôle selon son profil et valide la géométrie et les types de données.</summary>
        /// <param name="child">Contrôle MSForms direct à lire.</param>
        /// <param name="profile">Profil positif du type du contrôle.</param>
        /// <param name="sourcePath">Chemin source utilisé dans les erreurs.</param>
        /// <param name="proposedPath">Chemin de destination calculé.</param>
        /// <param name="newName">Nom proposé pour le contrôle copié.</param>
        /// <returns>Snapshot de propriétés autorisées pour la copie.</returns>
        /// <exception cref="InvalidOperationException">Une propriété ne respecte pas les limites du profil positif.</exception>
        private static ProfiledChildSnapshot ReadProfiledChild(dynamic child,
            VerifiedChildProfile profile, string sourcePath, string proposedPath, string newName)
        {
            var item = new ProfiledChildSnapshot { Profile = profile, SourcePath = sourcePath,
                ProposedName = newName, ProposedPath = proposedPath,
                Left = Convert.ToDouble(child.Left, CultureInfo.InvariantCulture),
                Top = Convert.ToDouble(child.Top, CultureInfo.InvariantCulture),
                Width = Convert.ToDouble(child.Width, CultureInfo.InvariantCulture),
                Height = Convert.ToDouble(child.Height, CultureInfo.InvariantCulture) };
            ValidateCopyBox(item.Left, item.Top, item.Width, item.Height, profile.Type);
            switch (profile.Type)
            {
                case "Label":
                    item.Caption = (string)child.Caption;
                    item.BackColor = CopyOleColor((object)child.BackColor);
                    item.FontName = (string)child.Font.Name;
                    item.FontSize = Convert.ToDouble(child.Font.Size, CultureInfo.InvariantCulture);
                    item.FontBold = (bool)child.Font.Bold;
                    if (string.IsNullOrWhiteSpace(item.FontName) || !IsFinite(item.FontSize) ||
                        item.FontSize <= 0 || item.FontSize > 200)
                        throw new InvalidOperationException("Label font is outside the tested range.");
                    break;
                case "TextBox":
                    object text = child.Value;
                    if (text != null && !(text is string))
                        throw new InvalidOperationException("TextBox.Value must be text or empty.");
                    item.TextValue = (string)text;
                    break;
                case "CheckBox":
                    item.Caption = (string)child.Caption;
                    object check = child.Value;
                    if (!(check is bool))
                        throw new InvalidOperationException("CheckBox.Value must be Boolean.");
                    item.BooleanValue = (bool)check;
                    break;
                case "CommandButton":
                case "OptionButton":
                    item.Caption = (string)child.Caption;
                    break;
                case "ComboBox":
                    object listWidth = child.ListWidth;
                    if (!(listWidth is string) || string.IsNullOrWhiteSpace((string)listWidth) ||
                        ((string)listWidth).Length > 64)
                        throw new InvalidOperationException("ComboBox.ListWidth must be nonempty text.");
                    item.ListWidth = (string)listWidth;
                    break;
                default: throw new InvalidOperationException("Profile has no copy implementation.");
            }
            return item;
        }

        /// <summary>Copie une Frame racine avec ses enfants directs profilés, relit chaque propriété autorisée et annule les ajouts en cas d’échec.</summary>
        /// <param name="request">Projet, formulaire, chemin source, version attendue et nom proposé.</param>
        /// <returns>Rapport partiel de duplication et arbre de contrôles relu.</returns>
        /// <exception cref="ArgumentException">Les paramètres requis ou le chemin racine sont invalides.</exception>
        /// <exception cref="InvalidOperationException">La prévalidation échoue, l’arbre change ou une vérification/compensation échoue.</exception>
        public object DuplicateFrameProfiled(Request request)
        {
            ProfiledFrameSnapshot plan = ReadFrameProfilePlan(request);
            if (plan.Issues.Count != 0)
                throw new InvalidOperationException("Frame profile copy plan is ineligible: " +
                    string.Join("; ", plan.Issues));
            dynamic form = GetForm(GetDesignProject(request.Project), request.Form);
            dynamic fresh = Tree(request.Project, request.Form);
            if (!string.Equals((string)fresh.TreeVersion, plan.TreeVersion,
                StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The UserForm hierarchy changed during copy preflight.");
            dynamic rootControls = form.Designer.Controls;
            dynamic copiedFrame = null;
            bool frameCreated = false;
            var createdChildren = new List<string>();
            try
            {
                copiedFrame = rootControls.Add("Forms.Frame.1", plan.NewName, true);
                frameCreated = true;
                copiedFrame.Left = plan.Left;
                copiedFrame.Top = plan.Top;
                copiedFrame.Width = plan.Width;
                copiedFrame.Height = plan.Height;
                copiedFrame.Caption = plan.Caption;
                foreach (ProfiledChildSnapshot item in plan.Children)
                {
                    dynamic copy = copiedFrame.Controls.Add(item.Profile.ProgId, item.ProposedName, true);
                    createdChildren.Add(item.ProposedName);
                    ApplyProfiledChild(copy, item);
                }
                dynamic after = Tree(request.Project, request.Form);
                string framePath = "Controls/" + plan.NewName;
                if (!TreeContainsPath((IEnumerable)after.Controls, framePath) ||
                    (int)after.NodeCount != plan.NodeCount + 1 + plan.Children.Count ||
                    string.Equals((string)after.TreeVersion, plan.TreeVersion, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("The new Frame hierarchy was not reflected in form_tree.");
                dynamic installedFrame = ResolveTreeItem(form.Designer, framePath);
                if ((string)installedFrame.Caption != plan.Caption ||
                    Convert.ToInt32(installedFrame.Controls.Count, CultureInfo.InvariantCulture) != plan.Children.Count ||
                    !SameCopyBox(installedFrame, plan.Left, plan.Top, plan.Width, plan.Height))
                    throw new InvalidOperationException("Copied Frame properties differ from the preflight snapshot.");
                foreach (ProfiledChildSnapshot item in plan.Children)
                {
                    if (!TreeContainsPath((IEnumerable)after.Controls, item.ProposedPath))
                        throw new InvalidOperationException("Copied child missing from form_tree: " + item.ProposedPath);
                    VerifyProfiledChild(ResolveTreeItem(form.Designer, item.ProposedPath), item);
                }
                return new { SourcePath = plan.SourcePath, NewPath = framePath,
                    DirectChildrenCopied = plan.Children.Count,
                    CopiedChildren = plan.Children.Select(item => new { Path = item.ProposedPath,
                        Type = item.Profile.Type, CopiedFields = item.Profile.Fields,
                        Limitation = item.Profile.Limitation }).ToArray(),
                    Completeness = "Partial", Tree = after };
            }
            catch (Exception originalError)
            {
                var rollbackErrors = new List<string>();
                if (frameCreated)
                {
                    for (int index = createdChildren.Count - 1; index >= 0; index--)
                        try { copiedFrame.Controls.Remove(createdChildren[index]); }
                        catch (Exception ex) { rollbackErrors.Add(createdChildren[index] + ": " + ex.Message); }
                    try { rootControls.Remove(plan.NewName); }
                    catch (Exception ex) { rollbackErrors.Add(plan.NewName + ": " + ex.Message); }
                }
                if (rollbackErrors.Count > 0)
                    throw new InvalidOperationException("Profiled Frame copy failed and rollback was incomplete. " +
                        string.Join("; ", rollbackErrors) + " Inspect form_tree before retrying.", originalError);
                throw;
            }
        }

        /// <summary>Applique uniquement les setters autorisés par le profil du contrôle copié.</summary>
        /// <param name="copy">Nouveau contrôle MSForms.</param>
        /// <param name="item">Snapshot des valeurs autorisées.</param>
        /// <exception cref="InvalidOperationException">Aucun setter positif n’existe pour le type du profil.</exception>
        private static void ApplyProfiledChild(dynamic copy, ProfiledChildSnapshot item)
        {
            copy.Left = item.Left;
            copy.Top = item.Top;
            copy.Width = item.Width;
            copy.Height = item.Height;
            switch (item.Profile.Type)
            {
                case "Label":
                    copy.Caption = item.Caption;
                    copy.BackColor = item.BackColor;
                    dynamic font = copy.Font;
                    font.Name = item.FontName;
                    font.Size = item.FontSize;
                    font.Bold = item.FontBold;
                    break;
                case "TextBox": if (item.TextValue != null) copy.Value = item.TextValue; break;
                case "CheckBox":
                    copy.Caption = item.Caption;
                    if (item.BooleanValue) copy.Value = true;
                    break;
                case "CommandButton":
                case "OptionButton": copy.Caption = item.Caption; break;
                case "ComboBox": copy.ListWidth = item.ListWidth; break;
                default: throw new InvalidOperationException("Profile has no setter implementation.");
            }
        }

        /// <summary>Compare au snapshot la géométrie et les propriétés prises en charge du contrôle relu.</summary>
        /// <param name="actual">Contrôle recréé et résolu dans le formulaire.</param>
        /// <param name="item">Valeurs source attendues et profil appliqué.</param>
        /// <exception cref="InvalidOperationException">Une propriété relue diffère ou le profil ne permet pas de lecture.</exception>
        private static void VerifyProfiledChild(dynamic actual, ProfiledChildSnapshot item)
        {
            if (!SameCopyBox(actual, item.Left, item.Top, item.Width, item.Height))
                throw new InvalidOperationException("Copied child geometry differs: " + item.ProposedPath);
            switch (item.Profile.Type)
            {
                case "Label":
                    if ((string)actual.Caption != item.Caption ||
                        CopyOleColor((object)actual.BackColor) != item.BackColor ||
                        (string)actual.Font.Name != item.FontName ||
                        Math.Abs(Convert.ToDouble(actual.Font.Size, CultureInfo.InvariantCulture) - item.FontSize) > 0.01 ||
                        (bool)actual.Font.Bold != item.FontBold)
                        throw new InvalidOperationException("Copied Label differs: " + item.ProposedPath);
                    break;
                case "TextBox":
                    if (!string.Equals(item.TextValue, actual.Value as string, StringComparison.Ordinal))
                        throw new InvalidOperationException("Copied TextBox value differs: " + item.ProposedPath);
                    break;
                case "CheckBox":
                    if ((string)actual.Caption != item.Caption ||
                        !(actual.Value is bool) || (bool)actual.Value != item.BooleanValue)
                        throw new InvalidOperationException("Copied CheckBox differs: " + item.ProposedPath);
                    break;
                case "CommandButton":
                case "OptionButton":
                    if ((string)actual.Caption != item.Caption)
                        throw new InvalidOperationException("Copied button caption differs: " + item.ProposedPath);
                    break;
                case "ComboBox":
                    if (!string.Equals(item.ListWidth, actual.ListWidth as string, StringComparison.Ordinal))
                        throw new InvalidOperationException("Copied ComboBox ListWidth differs: " + item.ProposedPath);
                    break;
                default: throw new InvalidOperationException("Profile has no readback implementation.");
            }
        }
    }
}
