using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;

namespace VBAi
{
    /// <summary>Implémente une copie transactionnelle limitée aux Frames racines avec Labels directs.</summary>
    internal sealed partial class VbeForms
    {
        /// <summary>Capture le chemin source et les propriétés de Label autorisées pour une copie de Frame.</summary>
        private sealed class DirectLabelCopy
        {
            /// <summary>Chemin canonique du Label source.</summary>
            public string SourcePath;
            /// <summary>Nom proposé pour le Label copié.</summary>
            public string Name;
            /// <summary>Légende du Label.</summary>
            public string Caption;
            /// <summary>Position horizontale dans la Frame.</summary>
            public double Left;
            /// <summary>Position verticale dans la Frame.</summary>
            public double Top;
            /// <summary>Largeur du Label.</summary>
            public double Width;
            /// <summary>Hauteur du Label.</summary>
            public double Height;
            /// <summary>Couleur OLE du fond du Label.</summary>
            public int BackColor;
            /// <summary>Nom de police du Label.</summary>
            public string FontName;
            /// <summary>Taille de police du Label.</summary>
            public double FontSize;
            /// <summary>Indique si la police du Label est en gras.</summary>
            public bool FontBold;
        }

        // Transactional probe for one root Frame with direct Label children.
        // Other child types, nested Frames and unqualified properties are refused.
        /// <summary>Duplique une Frame racine et ses Labels directs admissibles, vérifie le résultat et annule les ajouts si une étape échoue.</summary>
        /// <param name="request">Projet, formulaire, chemin racine, version attendue et nom proposé.</param>
        /// <returns>Rapport partiel avec la nouvelle Frame, les Labels copiés et l’arbre relu.</returns>
        /// <exception cref="ArgumentException">Le chemin source ne désigne pas une Frame racine valide.</exception>
        /// <exception cref="InvalidOperationException">Le plan est inadmissible, la version est périmée, une vérification échoue ou le rollback est incomplet.</exception>
        public object DuplicateFrameWithLabels(Request request)
        {
            if (string.IsNullOrWhiteSpace(request.ControlPath) ||
                request.ControlPath.Split('/').Length != 2)
                throw new ArgumentException("Only a root Controls/<Frame> path is supported by this probe.");
            dynamic plan = FrameCopyPlan(request);
            if (!(bool)plan.EligibleForLimitedProbe)
                throw new InvalidOperationException("Frame copy plan is ineligible: " +
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

            var labels = new List<DirectLabelCopy>();
            foreach (dynamic childPlan in plan.Children)
            {
                dynamic label = ResolveTreeItem(form.Designer, (string)childPlan.SourcePath);
                var item = new DirectLabelCopy {
                    SourcePath = (string)childPlan.SourcePath,
                    Name = (string)childPlan.ProposedName,
                    Caption = (string)label.Caption,
                    Left = Convert.ToDouble(label.Left, CultureInfo.InvariantCulture),
                    Top = Convert.ToDouble(label.Top, CultureInfo.InvariantCulture),
                    Width = Convert.ToDouble(label.Width, CultureInfo.InvariantCulture),
                    Height = Convert.ToDouble(label.Height, CultureInfo.InvariantCulture),
                    BackColor = CopyOleColor((object)label.BackColor),
                    FontName = (string)label.Font.Name,
                    FontSize = Convert.ToDouble(label.Font.Size, CultureInfo.InvariantCulture),
                    FontBold = (bool)label.Font.Bold
                };
                ValidateCopyBox(item.Left, item.Top, item.Width, item.Height, "Label " + item.Name);
                if (string.IsNullOrWhiteSpace(item.FontName) || !IsFinite(item.FontSize) ||
                    item.FontSize <= 0 || item.FontSize > 200)
                    throw new InvalidOperationException("Source Label has an unsupported font: " + item.SourcePath);
                labels.Add(item);
            }
            if (labels.Count == 0)
                throw new InvalidOperationException("Use duplicate_empty_form_frame for an empty Frame.");

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
                foreach (DirectLabelCopy item in labels)
                {
                    dynamic copy = copiedFrame.Controls.Add("Forms.Label.1", item.Name, true);
                    createdChildren.Add(item.Name);
                    copy.Left = item.Left;
                    copy.Top = item.Top;
                    copy.Width = item.Width;
                    copy.Height = item.Height;
                    copy.Caption = item.Caption;
                    copy.BackColor = item.BackColor;
                    dynamic font = copy.Font;
                    font.Name = item.FontName;
                    font.Size = item.FontSize;
                    font.Bold = item.FontBold;
                }

                dynamic after = Tree(request.Project, request.Form);
                string framePath = "Controls/" + request.NewName;
                if (!TreeContainsPath((IEnumerable)after.Controls, framePath) ||
                    (int)after.NodeCount != (int)before.NodeCount + 1 + labels.Count ||
                    string.Equals((string)after.TreeVersion, request.ExpectedTreeVersion,
                        StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("The Frame and Labels were not reflected in form_tree.");
                dynamic installedFrame = ResolveTreeItem(form.Designer, framePath);
                if ((string)installedFrame.Caption != frameCaption ||
                    Convert.ToInt32(installedFrame.Controls.Count, CultureInfo.InvariantCulture) != labels.Count ||
                    !SameCopyBox(installedFrame, frameLeft, frameTop, frameWidth, frameHeight))
                    throw new InvalidOperationException("The copied Frame did not retain its supported properties.");
                foreach (DirectLabelCopy item in labels)
                {
                    string childPath = framePath + "/Controls/" + item.Name;
                    if (!TreeContainsPath((IEnumerable)after.Controls, childPath))
                        throw new InvalidOperationException("Copied Label is absent from form_tree: " + childPath);
                    dynamic installed = ResolveTreeItem(form.Designer, childPath);
                    if ((string)installed.Caption != item.Caption ||
                        !SameCopyBox(installed, item.Left, item.Top, item.Width, item.Height) ||
                        CopyOleColor((object)installed.BackColor) != item.BackColor ||
                        (string)installed.Font.Name != item.FontName ||
                        Math.Abs(Convert.ToDouble(installed.Font.Size, CultureInfo.InvariantCulture) - item.FontSize) > 0.01 ||
                        (bool)installed.Font.Bold != item.FontBold)
                        throw new InvalidOperationException("Copied Label differs from source: " + childPath);
                }
                return new { SourcePath = request.ControlPath, NewPath = framePath,
                    DirectLabelsCopied = labels.Count,
                    CopiedChildPaths = labels.Select(item => framePath + "/Controls/" + item.Name).ToArray(),
                    FrameProperties = new[] { "Name", "Caption", "Left", "Top", "Width", "Height" },
                    LabelProperties = new[] { "Name", "Caption", "Left", "Top", "Width", "Height",
                        "BackColor", "Font.Name", "Font.Size", "Font.Bold" },
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
                    throw new InvalidOperationException("Frame copy failed and rollback was incomplete. " +
                        string.Join("; ", rollbackErrors) + " Inspect form_tree before retrying.", originalError);
                throw;
            }
        }

        /// <summary>Convertit une couleur Drawing en entier OLE ou convertit directement la valeur numérique.</summary>
        /// <param name="value">Valeur couleur exposée par COM.</param>
        /// <returns>Entier de couleur au format OLE.</returns>
        private static int CopyOleColor(object value)
        {
            return value is Color ? ColorTranslator.ToOle((Color)value) :
                Convert.ToInt32(value, CultureInfo.InvariantCulture);
        }

        /// <summary>Vérifie que la géométrie est finie, positive et comprise dans la plage prise en charge.</summary>
        /// <param name="left">Coordonnée horizontale.</param>
        /// <param name="top">Coordonnée verticale.</param>
        /// <param name="width">Largeur positive.</param>
        /// <param name="height">Hauteur positive.</param>
        /// <param name="label">Nom utilisé pour identifier une erreur de géométrie.</param>
        /// <exception cref="InvalidOperationException">Une dimension est hors de la plage prise en charge.</exception>
        private static void ValidateCopyBox(double left, double top, double width, double height, string label)
        {
            if (!IsFinite(left) || !IsFinite(top) || !IsFinite(width) || !IsFinite(height) ||
                left < 0 || top < 0 || width <= 0 || height <= 0 ||
                left > 32767 || top > 32767 || width > 32767 || height > 32767)
                throw new InvalidOperationException(label + " geometry is outside the supported range.");
        }

        /// <summary>Compare la géométrie COM relue à la géométrie attendue avec une tolérance de 0,01.</summary>
        /// <param name="control">Contrôle COM à comparer.</param>
        /// <param name="left">Coordonnée horizontale attendue.</param>
        /// <param name="top">Coordonnée verticale attendue.</param>
        /// <param name="width">Largeur attendue.</param>
        /// <param name="height">Hauteur attendue.</param>
        /// <returns>Vrai si les quatre mesures restent dans la tolérance.</returns>
        private static bool SameCopyBox(dynamic control, double left, double top, double width, double height)
        {
            return Math.Abs(Convert.ToDouble(control.Left, CultureInfo.InvariantCulture) - left) <= 0.01 &&
                Math.Abs(Convert.ToDouble(control.Top, CultureInfo.InvariantCulture) - top) <= 0.01 &&
                Math.Abs(Convert.ToDouble(control.Width, CultureInfo.InvariantCulture) - width) <= 0.01 &&
                Math.Abs(Convert.ToDouble(control.Height, CultureInfo.InvariantCulture) - height) <= 0.01;
        }
    }
}
