using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace CodexVBE
{
    /// <summary>Aligne, espace et ordonne les contrôles directs d’un conteneur UserForm.</summary>
    internal sealed partial class VbeForms
    {
        /// <summary>Calcule ou applique un plan de géométrie pour des contrôles sélectionnés dans un même conteneur.</summary>
        /// <param name="request">Chemins canoniques, action, taille de pas et version d’arbre attendue.</param>
        /// <param name="preview">Si true, retourne le plan sans modifier le Designer.</param>
        /// <returns>Plan de disposition en prévisualisation ou état vérifié après application.</returns>
        /// <exception cref="ArgumentException">Les contrôles ne sont pas uniques, canoniques ou dans un même conteneur.</exception>
        /// <exception cref="InvalidOperationException">L’arbre a changé ou le résultat ne peut pas être appliqué sûrement.</exception>
        public object LayoutControls(Request request, bool preview)
        {
            if (request.Items == null || request.Items.Length < 2 || request.Items.Length > 64 ||
                request.Items.Distinct(StringComparer.OrdinalIgnoreCase).Count() != request.Items.Length ||
                string.IsNullOrWhiteSpace(request.ExpectedTreeVersion))
                throw new ArgumentException("Items must contain 2 to 64 distinct canonical control paths, with ExpectedTreeVersion.");
            dynamic form = GetForm(GetDesignProject(request.Project), request.Form);
            dynamic tree = Tree(request.Project, request.Form);
            if (!string.Equals((string)tree.TreeVersion, request.ExpectedTreeVersion, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The form hierarchy changed since it was read.");
            string parent = null;
            var controls = new List<object>(); var boxes = new List<FormLayoutBox>();
            foreach (string path in request.Items)
            {
                if (path == null || !TreeContainsPath((IEnumerable)tree.Controls, path)) throw new ArgumentException("Every path must exist in form_tree.");
                int end = path.LastIndexOf("Controls/", StringComparison.Ordinal);
                // Every canonical tree path starts with Controls/; Pages and Tabs remain rejected here.
                if (path.Substring(end + 9).Contains("/")) throw new ArgumentException("Select controls, not Pages or Tabs.");
                string owner = path.Substring(0, end).TrimEnd('/');
                if (parent != null && parent != owner) throw new ArgumentException("All selected controls must share one container.");
                parent = owner;
                dynamic control = ResolveTreeItem(form.Designer, path);
                controls.Add((object)control);
                boxes.Add(new FormLayoutBox { Path = path, Left = (double)control.Left, Top = (double)control.Top, Width = (double)control.Width, Height = (double)control.Height });
            }
            dynamic container = parent.Length == 0 ? form.Designer : ResolveTreeItem(form.Designer, parent);
            double width = (double)container.InsideWidth, height = (double)container.InsideHeight;
            var plan = FormLayoutPlan.Create(boxes.ToArray(), request.Action, request.Width, width, height);
            if (preview) return new { request.Form, Before = boxes.ToArray(), After = plan, request.ExpectedTreeVersion };
            int started = 0;
            try
            {
                for (int i = 0; i < controls.Count; i++) { started = i + 1; ApplyBox(controls[i], plan[i]); }
                for (int i = 0; i < controls.Count; i++) VerifyBox(controls[i], plan[i]);
            }
            catch (Exception original)
            {
                var errors = new List<string>();
                for (int i = started - 1; i >= 0; i--) try { ApplyBox(controls[i], boxes[i]); VerifyBox(controls[i], boxes[i]); } catch (Exception ex) { errors.Add(ex.Message); }
                throw new InvalidOperationException(original.Message + (errors.Count == 0 ? " Original geometry restored." : " Geometry rollback incomplete: " + string.Join("; ", errors)), original);
            }
            return new { Applied = true, Verified = true, Saved = false, Tree = Tree(request.Project, request.Form) };
        }
        /// <summary>Écrit la taille puis la position prévues sur un contrôle Designer.</summary>
        /// <param name="control">Contrôle natif à déplacer et redimensionner.</param>
        /// <param name="box">Géométrie cible.</param>
        private static void ApplyBox(dynamic control, FormLayoutBox box)
        { control.Width = box.Width; control.Height = box.Height; control.Left = box.Left; control.Top = box.Top; }
        /// <summary>Relit la géométrie native et vérifie qu’elle correspond à la boîte demandée.</summary>
        /// <param name="control">Contrôle à vérifier.</param>
        /// <param name="box">Géométrie attendue.</param>
        /// <exception cref="InvalidOperationException">Le Designer n’a pas conservé les dimensions ou coordonnées.</exception>
        private static void VerifyBox(dynamic control, FormLayoutBox box)
        {
            if (new[] { (double)control.Left, (double)control.Top, (double)control.Width, (double)control.Height }.Any(x => double.IsNaN(x) || double.IsInfinity(x)) || Math.Abs((double)control.Left - box.Left) > 0.1 || Math.Abs((double)control.Top - box.Top) > 0.1 || Math.Abs((double)control.Width - box.Width) > 0.1 || Math.Abs((double)control.Height - box.Height) > 0.1)
                throw new InvalidOperationException("The designer did not retain the requested geometry for " + box.Path);
        }

        /// <summary>Réordonne les indices de tabulation des contrôles directs d’un conteneur.</summary>
        /// <param name="request">Ordre complet des noms, parent facultatif et version d’arbre attendue.</param>
        /// <returns>Arbre relu après application et vérification de l’ordre.</returns>
        /// <exception cref="ArgumentException">La liste n’énumère pas tous les contrôles directs exactement une fois.</exception>
        /// <exception cref="InvalidOperationException">L’arbre a changé ou l’application et son retour arrière ont échoué.</exception>
        public object SetTabOrder(Request request)
        {
            if (request.Items == null || request.Items.Length == 0 || request.Items.Length > 64 ||
                request.Items.Distinct(StringComparer.OrdinalIgnoreCase).Count() != request.Items.Length)
                throw new ArgumentException("Items must list every direct control name in the desired tab order.");
            dynamic form = GetForm(GetDesignProject(request.Project), request.Form);
            dynamic tree = Tree(request.Project, request.Form);
            if (string.IsNullOrEmpty(request.ExpectedTreeVersion) || !string.Equals((string)tree.TreeVersion, request.ExpectedTreeVersion, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The form hierarchy changed since it was read.");
            if (!string.IsNullOrEmpty(request.ParentPath) && !TreeContainsPath((IEnumerable)tree.Controls, request.ParentPath))
                throw new ArgumentException("ParentPath is not canonical in form_tree.");
            dynamic parent = string.IsNullOrEmpty(request.ParentPath) ? form.Designer : ResolveTreeItem(form.Designer, request.ParentPath);
            var items = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
            var oldOrder = new SortedDictionary<int, object>();
            foreach (dynamic control in parent.Controls) { items.Add((string)control.Name, (object)control); oldOrder.Add((int)control.TabIndex, (object)control); }
            if (items.Count != request.Items.Length || request.Items.Any(x => x == null || !items.ContainsKey(x)))
                throw new ArgumentException("Items must contain all direct controls exactly once.");
            try
            {
                for (int i = request.Items.Length - 1; i >= 0; i--) ((dynamic)items[request.Items[i]]).TabIndex = 0;
                for (int i = 0; i < request.Items.Length; i++) if ((int)((dynamic)items[request.Items[i]]).TabIndex != i) throw new InvalidOperationException("Tab order readback mismatch.");
            }
            catch (Exception error)
            {
                try {
                    foreach (var item in oldOrder.Reverse()) ((dynamic)item.Value).TabIndex = 0;
                    foreach (var item in oldOrder) if ((int)((dynamic)item.Value).TabIndex != item.Key) throw new InvalidOperationException("Tab order rollback readback mismatch.");
                } catch (Exception rollback) { throw new InvalidOperationException(error.Message + " Rollback failed: " + rollback.Message, error); }
                throw new InvalidOperationException(error.Message + " Original tab order restored.", error);
            }
            return new { Applied = true, Verified = true, Saved = false, Tree = Tree(request.Project, request.Form) };
        }
    }
}
