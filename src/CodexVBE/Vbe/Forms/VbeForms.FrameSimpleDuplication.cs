using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;

namespace CodexVBE
{
    /// <summary>Implémente la copie partielle d’une Frame et des Labels/TextBox directs de son profil simple.</summary>
    internal sealed partial class VbeForms
    {
        /// <summary>Capture les valeurs autorisées à recopier sur un Label ou TextBox enfant.</summary>
        private sealed class SimpleFrameChild
        {
            /// <summary>Type MSForms du contrôle enfant.</summary>
            public string Kind;
            /// <summary>Nom proposé pour le contrôle copié.</summary>
            public string Name;
            /// <summary>Légende copiée pour un Label.</summary>
            public string Caption;
            /// <summary>Valeur copiée pour un TextBox lorsque sa source contient du texte.</summary>
            public string TextValue;
            /// <summary>Position horizontale de l’enfant dans sa Frame.</summary>
            public double Left;
            /// <summary>Position verticale de l’enfant dans sa Frame.</summary>
            public double Top;
            /// <summary>Largeur de l’enfant.</summary>
            public double Width;
            /// <summary>Hauteur de l’enfant.</summary>
            public double Height;
            /// <summary>Couleur OLE convertie du fond du Label.</summary>
            public int BackColor;
            /// <summary>Nom de police copié pour le Label.</summary>
            public string FontName;
            /// <summary>Taille de police copiée pour le Label.</summary>
            public double FontSize;
            /// <summary>Indique si la police du Label est en gras.</summary>
            public bool FontBold;
        }

        // Extends the proven Frame+Labels transaction with the narrow TextBox
        // profile. It is a separate bridge route so the Label-only contract
        // and its tested rejection of TextBox children stay unchanged.
        /// <summary>Copie une Frame racine avec ses enfants Label et TextBox directs admissibles puis vérifie la hiérarchie, la géométrie et les propriétés prévues.</summary>
        /// <param name="request">Projet, formulaire, chemin racine, version attendue et nom de la nouvelle Frame.</param>
        /// <returns>Rapport partiel avec les chemins des enfants copiés et le nouvel arbre.</returns>
        /// <exception cref="ArgumentException">Le chemin ou les données obligatoires ne conviennent pas au profil racine.</exception>
        /// <exception cref="InvalidOperationException">Le plan est inadmissible, un enfant est hors profil, la vérification échoue ou le rollback est incomplet.</exception>
        public object DuplicateFrameWithSimpleChildren(Request request)
        {
            if (string.IsNullOrWhiteSpace(request.ControlPath) ||
                request.ControlPath.Split('/').Length != 2)
                throw new ArgumentException("Only a root Controls/<Frame> path is supported by this probe.");
            dynamic plan = FrameSimpleCopyPlan(request);
            if (!(bool)plan.EligibleForLimitedProbe)
                throw new InvalidOperationException("Frame simple copy plan is ineligible: " +
                    string.Join("; ", (IEnumerable<string>)plan.Issues));
            dynamic form = GetForm(GetDesignProject(request.Project), request.Form);
            dynamic before = Tree(request.Project, request.Form);
            if (!string.Equals((string)before.TreeVersion, request.ExpectedTreeVersion,
                StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The UserForm hierarchy changed since it was read.");
            dynamic source = ResolveTreeItem(form.Designer, request.ControlPath);
            string frameCaption = (string)source.Caption;
            double frameLeft = Convert.ToDouble(source.Left, CultureInfo.InvariantCulture);
            double frameTop = Convert.ToDouble(source.Top, CultureInfo.InvariantCulture);
            double frameWidth = Convert.ToDouble(source.Width, CultureInfo.InvariantCulture);
            double frameHeight = Convert.ToDouble(source.Height, CultureInfo.InvariantCulture);
            ValidateCopyBox(frameLeft, frameTop, frameWidth, frameHeight, "Frame");

            var children = new List<SimpleFrameChild>();
            int textBoxCount = 0;
            foreach (dynamic childPlan in plan.Children)
            {
                dynamic control = ResolveTreeItem(form.Designer, (string)childPlan.SourcePath);
                string kind = (string)childPlan.Type;
                var item = new SimpleFrameChild {
                    Kind = kind,
                    Name = (string)childPlan.ProposedName,
                    Left = Convert.ToDouble(control.Left, CultureInfo.InvariantCulture),
                    Top = Convert.ToDouble(control.Top, CultureInfo.InvariantCulture),
                    Width = Convert.ToDouble(control.Width, CultureInfo.InvariantCulture),
                    Height = Convert.ToDouble(control.Height, CultureInfo.InvariantCulture)
                };
                ValidateCopyBox(item.Left, item.Top, item.Width, item.Height, kind + " " + item.Name);
                if (string.Equals(kind, "Label", StringComparison.OrdinalIgnoreCase))
                {
                    item.Caption = (string)control.Caption;
                    item.BackColor = CopyOleColor((object)control.BackColor);
                    item.FontName = (string)control.Font.Name;
                    item.FontSize = Convert.ToDouble(control.Font.Size, CultureInfo.InvariantCulture);
                    item.FontBold = (bool)control.Font.Bold;
                    if (string.IsNullOrWhiteSpace(item.FontName) || !IsFinite(item.FontSize) ||
                        item.FontSize <= 0 || item.FontSize > 200)
                        throw new InvalidOperationException("Source Label has an unsupported font: " + childPlan.SourcePath);
                }
                else
                {
                    object raw = control.Value;
                    if (raw != null && !(raw is string))
                        throw new InvalidOperationException("TextBox.Value must be text or empty: " + childPlan.SourcePath);
                    item.TextValue = (string)raw;
                    textBoxCount++;
                }
                children.Add(item);
            }
            if (textBoxCount == 0)
                throw new InvalidOperationException("Use duplicate_form_frame_labels when no TextBox child exists.");

            dynamic rootControls = form.Designer.Controls;
            dynamic copiedFrame = null;
            bool frameCreated = false;
            var createdChildren = new List<string>();
            try
            {
                copiedFrame = rootControls.Add("Forms.Frame.1", request.NewName, true);
                frameCreated = true;
                copiedFrame.Left = frameLeft;
                copiedFrame.Top = frameTop;
                copiedFrame.Width = frameWidth;
                copiedFrame.Height = frameHeight;
                copiedFrame.Caption = frameCaption;
                foreach (SimpleFrameChild item in children)
                {
                    string progId = string.Equals(item.Kind, "Label", StringComparison.OrdinalIgnoreCase)
                        ? "Forms.Label.1" : "Forms.TextBox.1";
                    dynamic copy = copiedFrame.Controls.Add(progId, item.Name, true);
                    createdChildren.Add(item.Name);
                    copy.Left = item.Left;
                    copy.Top = item.Top;
                    copy.Width = item.Width;
                    copy.Height = item.Height;
                    if (progId == "Forms.Label.1")
                    {
                        copy.Caption = item.Caption;
                        copy.BackColor = item.BackColor;
                        dynamic font = copy.Font;
                        font.Name = item.FontName;
                        font.Size = item.FontSize;
                        font.Bold = item.FontBold;
                    }
                    else if (item.TextValue != null) copy.Value = item.TextValue;
                }

                dynamic after = Tree(request.Project, request.Form);
                string framePath = "Controls/" + request.NewName;
                if (!TreeContainsPath((IEnumerable)after.Controls, framePath) ||
                    (int)after.NodeCount != (int)before.NodeCount + 1 + children.Count ||
                    string.Equals((string)after.TreeVersion, request.ExpectedTreeVersion,
                        StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("The Frame and children were not reflected in form_tree.");
                dynamic installedFrame = ResolveTreeItem(form.Designer, framePath);
                if ((string)installedFrame.Caption != frameCaption ||
                    Convert.ToInt32(installedFrame.Controls.Count, CultureInfo.InvariantCulture) != children.Count ||
                    !SameCopyBox(installedFrame, frameLeft, frameTop, frameWidth, frameHeight))
                    throw new InvalidOperationException("The copied Frame did not retain its supported properties.");
                foreach (SimpleFrameChild item in children)
                {
                    string childPath = framePath + "/Controls/" + item.Name;
                    if (!TreeContainsPath((IEnumerable)after.Controls, childPath))
                        throw new InvalidOperationException("Copied child is absent from form_tree: " + childPath);
                    dynamic installed = ResolveTreeItem(form.Designer, childPath);
                    if (!SameCopyBox(installed, item.Left, item.Top, item.Width, item.Height))
                        throw new InvalidOperationException("Copied child geometry differs: " + childPath);
                    if (string.Equals(item.Kind, "Label", StringComparison.OrdinalIgnoreCase))
                    {
                        if ((string)installed.Caption != item.Caption ||
                            CopyOleColor((object)installed.BackColor) != item.BackColor ||
                            (string)installed.Font.Name != item.FontName ||
                            Math.Abs(Convert.ToDouble(installed.Font.Size, CultureInfo.InvariantCulture) - item.FontSize) > 0.01 ||
                            (bool)installed.Font.Bold != item.FontBold)
                            throw new InvalidOperationException("Copied Label differs: " + childPath);
                    }
                    else if (!string.Equals(item.TextValue, installed.Value as string, StringComparison.Ordinal))
                        throw new InvalidOperationException("Copied TextBox value differs: " + childPath);
                }
                return new { SourcePath = request.ControlPath, NewPath = framePath,
                    DirectLabelsCopied = children.Count - textBoxCount, DirectTextBoxesCopied = textBoxCount,
                    CopiedChildPaths = children.Select(item => framePath + "/Controls/" + item.Name).ToArray(),
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
                    try { rootControls.Remove(request.NewName); }
                    catch (Exception ex) { rollbackErrors.Add(request.NewName + ": " + ex.Message); }
                }
                if (rollbackErrors.Count > 0)
                    throw new InvalidOperationException("Frame simple copy failed and rollback was incomplete. " +
                        string.Join("; ", rollbackErrors) + " Inspect form_tree before retrying.", originalError);
                throw;
            }
        }
    }
}
