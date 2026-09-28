using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Web.Script.Serialization;

namespace CodexVBE
{
    /// <summary>Mesure et redimensionne les conteneurs UserForm selon les contrôles enfants visibles.</summary>
internal sealed partial class VbeForms
    {
        /// <summary>Measures the owning window DPI; replaceable by a bounded native contract probe.</summary>
        internal static Func<IntPtr, uint> MeasureFitWindowDpi = FitWindowDpi;
        /// <summary>Prévisualise les dimensions natives nécessaires pour contenir les enfants directs.</summary>
        /// <param name="request">Projet, formulaire, chemin canonique, version d'arbre et padding Left/Top.</param>
        /// <returns>Plan en lecture seule calculé à partir des mesures observées du concepteur.</returns>
        public object PreviewFitFormContent(Request request) { return FitFormContent(request, true); }

        /// <summary>Applique le plan de dimensionnement et vérifie les dimensions et les enfants relus.</summary>
        /// <param name="request">Action fit_container ou fit_scroll_extent et paramètres du plan.</param>
        /// <returns>Résultat vérifié ou incertitude après le début des écritures natives.</returns>
        public object ApplyFitFormContent(Request request) { return FitFormContent(request, false); }

                /// <summary>Valide une capture d’arbre, mesure les enfants, calcule les dimensions requises et peut les appliquer puis les relire.</summary>
        /// <param name="request">Conteneur, action, révision attendue et padding demandé.</param>
        /// <param name="preview">Si true, renvoie uniquement le plan sans écrire les dimensions.</param>
        /// <returns>Plan en prévisualisation ou résultat de mutation avec mesures relues et incertitude explicite.</returns>
        /// <exception cref="ArgumentException">La requête, l’action, le padding ou le chemin du conteneur est invalide.</exception>
        /// <exception cref="InvalidOperationException">L’arbre a changé ou les mesures natives et géométries ne permettent pas un ajustement sûr.</exception>
private object FitFormContent(Request request, bool preview)
        {
            if (request == null || string.IsNullOrWhiteSpace(request.ExpectedTreeVersion))
                throw new ArgumentException("ExpectedTreeVersion from form_tree is required.");
            if (request.Action != "fit_container" && request.Action != "fit_scroll_extent")
                throw new ArgumentException("Action must be fit_container or fit_scroll_extent.");
            if (!IsFinite(request.Left) || !IsFinite(request.Top) || request.Left < 0 ||
                request.Top < 0 || request.Left > 1000 || request.Top > 1000)
                throw new ArgumentException("Left/Top padding must be finite and between 0 and 1000 points.");
            if (!string.IsNullOrEmpty(request.ControlPath) && !string.IsNullOrEmpty(request.ParentPath))
                throw new ArgumentException("Specify ControlPath or ParentPath, not both.");
            dynamic form = GetForm(GetDesignProject(request.Project), request.Form);
            dynamic tree = Tree(request.Project, request.Form);
            if (!string.Equals((string)tree.TreeVersion, request.ExpectedTreeVersion, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The form hierarchy changed since it was read.");
            string path = string.IsNullOrEmpty(request.ControlPath) ? request.ParentPath ?? "" : request.ControlPath;
            bool root = path.Length == 0 || path == "UserForm";
            if (!root && !TreeContainsPath((IEnumerable)tree.Controls, path))
                throw new ArgumentException("The container path must exist exactly in form_tree.");
            object container = root ? (object)form.Designer : ResolveTreeItem(form.Designer, path);
            string type = TypeDescriptor.GetClassName(container);
            if ((root && !string.Equals(type, "UserForm", StringComparison.OrdinalIgnoreCase)) ||
                (!root && !new[] { "Frame", "Page", "MultiPage" }.Contains(type, StringComparer.OrdinalIgnoreCase)))
                throw new InvalidOperationException("Only native UserForm, Frame, Page or MultiPage containers are supported.");
            var children = ReadFitChildren(container, (string)form.Name);
            if (children.Count == 0) throw new InvalidOperationException("The container has no direct controls to fit.");
            double right = children.Max(item => item.Left + item.Width) + request.Left;
            double bottom = children.Max(item => item.Top + item.Height) + request.Top;
            bool scroll = request.Action == "fit_scroll_extent";
            string horizontal = scroll ? "ScrollWidth" : "Width";
            string vertical = scroll ? "ScrollHeight" : "Height";
            var x = RequireFitProperty(container, horizontal, true, root ? (object)form : null);
            var y = RequireFitProperty(container, vertical, true, root ? (object)form : null);
            double beforeX = ReadFitNumber(container, x), beforeY = ReadFitNumber(container, y);
            double insideX = ReadFitNumber(container, RequireFitProperty(container, "InsideWidth", false));
            double insideY = ReadFitNumber(container, RequireFitProperty(container, "InsideHeight", false));
            if (insideX <= 0 || insideY <= 0 || (!scroll && (beforeX < insideX || beforeY < insideY)))
                throw new InvalidOperationException("Native outer and inside measurements are inconsistent.");
            double afterX = scroll ? Math.Max(right, insideX) : right + beforeX - insideX;
            double afterY = scroll ? Math.Max(bottom, insideY) : bottom + beforeY - insideY;
            double tolerance = 0.1;
            if (root && NativeDesignerObject(container))
            {
                uint dpi = MeasureFitWindowDpi(new IntPtr(Convert.ToInt64(vbe.MainWindow.HWnd)));
                if (dpi < 48 || dpi > 768) throw new InvalidOperationException("Native designer DPI is unavailable.");
                tolerance = 72d / dpi;
                // Round upward to a native pixel: a fractional border must not
                // cause the requested content extent to be clipped.
                if (!scroll)
                {
                    afterX = Math.Ceiling(afterX / tolerance) * tolerance;
                    afterY = Math.Ceiling(afterY / tolerance) * tolerance;
                }
            }
            if (!IsFinite(afterX) || !IsFinite(afterY) || afterX <= 0 || afterY <= 0 ||
                afterX > 32767 || afterY > 32767)
                throw new InvalidOperationException("The proposed native dimensions are outside supported point bounds.");
            var plan = new { Project = request.Project, Form = request.Form, ContainerPath = root ? "UserForm" : path,
                Action = request.Action, HorizontalProperty = horizontal, VerticalProperty = vertical,
                WidthBefore = beforeX, HeightBefore = beforeY, WidthAfter = afterX, HeightAfter = afterY,
                InsideWidth = insideX, InsideHeight = insideY, PaddingHorizontal = request.Left, PaddingVertical = request.Top,
                Children = children, ExpectedTreeVersion = request.ExpectedTreeVersion, TolerancePoints = tolerance, RuntimeQualified = false };
            if (preview) return new { ReadOnly = true, Plan = plan };
            try
            {
                x.SetValue(container, Convert.ChangeType(afterX, x.PropertyType, CultureInfo.InvariantCulture));
                y.SetValue(container, Convert.ChangeType(afterY, y.PropertyType, CultureInfo.InvariantCulture));
                if (Math.Abs(ReadFitNumber(container, x) - afterX) > tolerance ||
                    Math.Abs(ReadFitNumber(container, y) - afterY) > tolerance)
                    throw new InvalidOperationException("Native dimension readback differs from the plan.");
                double actualInsideX = ReadFitNumber(container, RequireFitProperty(container, "InsideWidth", false));
                double actualInsideY = ReadFitNumber(container, RequireFitProperty(container, "InsideHeight", false));
                if (!scroll && (actualInsideX < right - 0.1 || actualInsideY < bottom - 0.1 ||
                    actualInsideX > right + tolerance || actualInsideY > bottom + tolerance))
                    throw new InvalidOperationException("Native inside dimensions differ after resizing; decoration cannot be inferred further.");
                if (scroll && (actualInsideX > afterX + 0.1 || actualInsideY > afterY + 0.1))
                    throw new InvalidOperationException("The resulting viewport exceeds the native scroll extent.");
                var serializer = new JavaScriptSerializer();
                if (serializer.Serialize(children) != serializer.Serialize(ReadFitChildren(container, (string)form.Name)))
                    throw new InvalidOperationException("Native resizing changed the observed child geometry.");
                return new { Applied = true, Verified = true, Uncertain = false, Saved = false, Plan = plan,
                    ActualInsideWidth = actualInsideX, ActualInsideHeight = actualInsideY,
                    Tree = Tree(request.Project, request.Form), RuntimeQualified = false };
            }
            catch (Exception error)
            {
                return new { Applied = false, Verified = false, Uncertain = true, MutationInvoked = true,
                    Reason = error.Message, Plan = plan, RetryAllowed = false,
                    Limit = "Inspect the designer before another operation. No retry or automatic rollback was performed." };
            }
        }

                /// <summary>Exige une propriété scalaire numérique native avant toute écriture.</summary>
        /// <param name="target">Conteneur dont la propriété doit être lue.</param>
        /// <param name="name">Nom de la propriété numérique recherchée.</param>
        /// <param name="write">Indique si le descripteur doit être modifiable.</param>
        /// <param name="component">VBComponent facultatif, source des dimensions Width/Height du formulaire racine.</param>
        /// <returns>Descripteur de propriété validé.</returns>
        /// <exception cref="InvalidOperationException">La propriété n’existe pas, n’est pas numérique ou ne permet pas l’écriture demandée.</exception>
private static PropertyDescriptor RequireFitProperty(object target, string name, bool write, object component = null)
        {
            var descriptor = TypeDescriptor.GetProperties(target).Find(name, false);
            // The native UserForm Designer omits Width/Height. VBIDE exposes those
            // design properties on VBComponent.Properties instead.
            if (descriptor == null && component != null && NativeDesignerObject(target) &&
                (name == "Width" || name == "Height"))
            {
                foreach (dynamic property in ((dynamic)component).Properties)
                    if ((string)property.Name == name && (int)property.NumIndices == 0)
                    {
                        object value = property.Value;
                        if (value != null) descriptor = new FitComponentProperty(name, (object)property, value.GetType());
                        break;
                    }
            }
            if (descriptor == null || (write && descriptor.IsReadOnly) ||
                !new[] { typeof(float), typeof(double), typeof(decimal), typeof(int), typeof(short) }.Contains(descriptor.PropertyType))
                throw new InvalidOperationException("Native numeric property is absent or not writable: " + name);
            return descriptor;
        }

                /// <summary>Lit le DPI propre à la fenêtre VBE pour la quantification des dimensions natives.</summary>
        /// <param name="window">Handle de la fenêtre hôte.</param>
        /// <returns>DPI effectif de cette fenêtre.</returns>
[DllImport("user32.dll", EntryPoint = "GetDpiForWindow")]
        private static extern uint FitWindowDpi(IntPtr window);

        /// <summary>Adapte les dimensions de conception VBIDE absentes du Designer MSForms natif.</summary>
        private sealed class FitComponentProperty : PropertyDescriptor
        {
            /// <summary>Propriété COM exposée par VBComponent.Properties.</summary>
private readonly object property;
            /// <summary>Type CLR de la valeur courante utilisée par le descripteur.</summary>
private readonly Type valueType;
            /// <summary>Crée un descripteur éditable pour une propriété de composant VBIDE.</summary>
            /// <param name="name">Nom de la propriété native.</param>
            /// <param name="property">Référence COM à la propriété.</param>
            /// <param name="valueType">Type de sa valeur actuelle.</param>
internal FitComponentProperty(string name, object property, Type valueType) : base(name, null)
            { this.property = property; this.valueType = valueType; }
            /// <summary>Obtient le type du composant cible du descripteur.</summary>
            /// <value>object, car la propriété est portée par une référence COM dynamique.</value>
public override Type ComponentType => typeof(object);
            /// <summary>Obtient le type CLR des valeurs de cette propriété.</summary>
            /// <value>Type déterminé à partir de sa valeur native courante.</value>
public override Type PropertyType => valueType;
            /// <summary>Indique que l’écriture est autorisée par ce descripteur.</summary>
            /// <value>Valeur false pour permettre la modification.</value>
public override bool IsReadOnly => false;
            /// <summary>Lit la valeur actuelle depuis la propriété COM.</summary>
            /// <param name="component">Composant cible, non utilisé par la référence déjà capturée.</param>
            /// <returns>Valeur native de la propriété.</returns>
public override object GetValue(object component) => ((dynamic)property).Value;
            /// <summary>Écrit une valeur convertie dans la propriété COM.</summary>
            /// <param name="component">Composant cible, non utilisé par la référence déjà capturée.</param>
            /// <param name="value">Nouvelle valeur native.</param>
public override void SetValue(object component, object value) { ((dynamic)property).Value = value; }
            /// <summary>Indique qu’aucune valeur par défaut distincte ne peut être restaurée.</summary>
            /// <param name="component">Composant cible.</param>
            /// <returns><see langword="false"/> car le descripteur ne fournit pas de valeur de réinitialisation.</returns>
public override bool CanResetValue(object component) => false;
            /// <summary>Refuse une réinitialisation implicite de la propriété de conception.</summary>
            /// <param name="component">Composant cible.</param>
            /// <exception cref="NotSupportedException">La réinitialisation n’est pas prise en charge.</exception>
public override void ResetValue(object component) { throw new NotSupportedException(); }
            /// <summary>Indique que le descripteur ne sérialise pas cette valeur.</summary>
            /// <param name="component">Composant cible.</param>
            /// <returns><see langword="false"/> pour éviter de sérialiser une valeur native temporaire.</returns>
public override bool ShouldSerializeValue(object component) => false;
        }

                /// <summary>Lit une mesure native et refuse les valeurs non finies ou négatives.</summary>
        /// <param name="target">Objet qui possède la propriété.</param>
        /// <param name="property">Descripteur de la mesure à lire.</param>
        /// <returns>Valeur convertie en points, bornée à 32767.</returns>
        /// <exception cref="InvalidOperationException">La valeur est nulle, non numérique, négative ou hors bornes.</exception>
private static double ReadFitNumber(object target, PropertyDescriptor property)
        {
            object value = property.GetValue(target);
            if (value == null) throw new InvalidOperationException("Native measurement is null: " + property.Name);
            double result = Convert.ToDouble(value, CultureInfo.InvariantCulture);
            if (!IsFinite(result) || result < 0 || result > 32767)
                throw new InvalidOperationException("Native measurement is not finite or is out of bounds: " + property.Name);
            return result;
        }

                /// <summary>Capture la géométrie des seuls enfants directs dont le type natif est connu.</summary>
        /// <param name="container">Conteneur dont la collection Controls sera énumérée.</param>
        /// <param name="formName">Nom du formulaire utilisé pour vérifier le parent de chaque contrôle.</param>
        /// <returns>Boîtes des contrôles directs avec leurs chemins et dimensions observées.</returns>
        /// <exception cref="InvalidOperationException">Le conteneur, un type de contrôle, son nom ou sa géométrie ne peut pas être vérifié.</exception>
private static List<FormLayoutBox> ReadFitChildren(object container, string formName)
        {
            var controls = TypeDescriptor.GetProperties(container).Find("Controls", false);
            object collection = controls == null ? null : controls.GetValue(container);
            if (!(collection is IEnumerable))
                throw new InvalidOperationException("This native container does not expose an enumerable Controls property.");
            var result = new List<FormLayoutBox>();
            var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (object child in (IEnumerable)collection)
            {
                dynamic node = child;
                if (!SameContainer((object)node.Parent, container, formName)) continue;
                string name = (string)node.Name;
                string type = TypeDescriptor.GetClassName(child);
                if (!BuiltInControls.Contains("Forms." + type + ".1") || !names.Add(name))
                    throw new InvalidOperationException("Unknown or ambiguous direct child control: " + name);
                if (result.Count >= 512) throw new InvalidOperationException("More than 512 direct controls.");
                result.Add(new FormLayoutBox { Path = name,
                    Left = ReadFitNumber(child, RequireFitProperty(child, "Left", false)),
                    Top = ReadFitNumber(child, RequireFitProperty(child, "Top", false)),
                    Width = ReadFitNumber(child, RequireFitProperty(child, "Width", false)),
                    Height = ReadFitNumber(child, RequireFitProperty(child, "Height", false)) });
            }
            return result;
        }
    }
}
