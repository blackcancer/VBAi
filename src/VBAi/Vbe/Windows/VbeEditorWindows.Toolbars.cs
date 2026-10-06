using System;
using System.Collections.Generic;
using System.Security.Cryptography;
using System.Text;
using System.Web.Script.Serialization;

namespace VBAi
{

    /// <summary>Lit et modifie les barres d’outils normales du VBE avec vérification des versions observées.</summary>
    internal sealed partial class VbeEditorWindows
    {

        /// <summary>Retourne l’état lisible des barres d’outils normales et leurs erreurs de lecture.</summary>
        /// <returns>Un objet sérialisable contenant les instantanés, les erreurs et la version de la collection.</returns>
        public object Toolbars()
        {
            var bars = new List<object>(); var errors = new List<string>();
            foreach (dynamic bar in vbe.CommandBars)
            {
                try { if ((int)bar.Type == 0) bars.Add(ToolbarSnapshot((object)bar)); }
                catch (Exception error) { errors.Add(error.Message); }
            }
            string collectionVersion = null;
            try { collectionVersion = ToolbarCollectionVersion(); } catch (Exception ex) { errors.Add(ex.Message); }
            return new { Toolbars = bars, Errors = errors, ProfileErrors = ToolbarProfileErrors.ToArray(), ToolbarCollectionVersion = collectionVersion,
                Scope = "Normal VBE command bars only; menu bars and shortcut menus are excluded." };
        }

        /// <summary>Capture les propriétés et la géométrie accessibles d’une barre, avec des empreintes distinctes.</summary>
        /// <param name="bar">Barre native à lire.</param>
        /// <returns>Un instantané sérialisable avec erreurs par propriété et versions calculées si la lecture est complète.</returns>
        private static object ToolbarSnapshot(dynamic bar)
        {
            var state = new Dictionary<string, object>(); var errors = new Dictionary<string, string>();
            Read(state, errors, "Name", () => (string)bar.Name);
            Read(state, errors, "Type", () => (int)bar.Type);
            Read(state, errors, "Visible", () => (bool)bar.Visible);
            Read(state, errors, "Enabled", () => (bool)bar.Enabled);
            Read(state, errors, "BuiltIn", () => (bool)bar.BuiltIn);
            Read(state, errors, "Protection", () => (int)bar.Protection);
            var geometry = new Dictionary<string, object>(); var geometryErrors = new Dictionary<string, string>();
            Read(geometry, geometryErrors, "Position", () => (int)bar.Position);
            Read(geometry, geometryErrors, "RowIndex", () => (int)bar.RowIndex);
            Read(geometry, geometryErrors, "Left", () => (int)bar.Left);
            Read(geometry, geometryErrors, "Top", () => (int)bar.Top);
            Read(geometry, geometryErrors, "Width", () => (int)bar.Width);
            Read(geometry, geometryErrors, "Height", () => (int)bar.Height);
            string version = null;
            if (errors.Count == 0)
                using (var sha = SHA256.Create())
                    version = BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(new JavaScriptSerializer().Serialize(state)))).Replace("-", "").ToLowerInvariant();
            string layoutVersion = null;
            if (version != null && geometryErrors.Count == 0)
                using (var sha = SHA256.Create())
                    layoutVersion = BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(new JavaScriptSerializer().Serialize(new { State = state, Geometry = geometry })))).Replace("-", "").ToLowerInvariant();
            return new { Properties = state, Errors = errors, Geometry = geometry, GeometryErrors = geometryErrors,
                ToolbarLayoutVersion = layoutVersion, WindowVersion = version, VersionScope = "Toolbar identity, visibility, enabled state and protection; geometry is observational." };
        }

        /// <summary>Affiche ou masque une barre après contrôle de l’empreinte de son état.</summary>
        /// <param name="request">Requête contenant l’action, le nom et la version attendue.</param>
        /// <returns>Le résultat de la mutation et les instantanés avant et après lecture.</returns>
        public object SetToolbarVisibility(Request request)
        {
            if (request.Action != "show" && request.Action != "hide") throw new ArgumentException("Action must be show or hide.");
            if (string.IsNullOrWhiteSpace(request.ObjectName) || string.IsNullOrWhiteSpace(request.ExpectedWindowVersion))
                throw new ArgumentException("ObjectName and ExpectedWindowVersion from list_toolbars are required.");
            object target = FindNormalToolbar(request.ObjectName);
            dynamic selected = target;
            dynamic before = ToolbarSnapshot(target);
            if (before.WindowVersion == null || !string.Equals((string)before.WindowVersion, request.ExpectedWindowVersion, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("Toolbar state changed or could not be read. Read list_toolbars again.");
            bool desired = request.Action == "show";
            if ((bool)selected.Visible == desired)
                return new { Applied = false, Verified = true, Before = (object)before, After = (object)before, PersistenceVerified = false };
            if (!desired && (((int)selected.Protection & 8) != 0)) throw new InvalidOperationException("The toolbar is protected against hiding.");
            if (desired && !(bool)selected.Enabled) throw new InvalidOperationException("The toolbar is disabled; showing it would require enabling it first.");
            string error = null; object after = null; bool? actual = null;
            try { selected.Visible = desired; SaveToolbarProfile(selected, false); }
            catch (Exception ex) { error = ex.Message; }
            try { after = ToolbarSnapshot(target); actual = (bool)selected.Visible; }
            catch (Exception ex) { error = error ?? ex.Message; }
            return new { Applied = actual.HasValue ? (bool?)(actual.Value != (bool)before.Properties["Visible"]) : null,
                Verified = error == null && actual == desired, Before = (object)before, After = after,
                NativeError = error, PersistenceVerified = false, NextRead = "list_toolbars" };
        }

        /// <summary>Résout une barre par son nom et refuse les menus, ambiguïtés et disparitions.</summary>
        /// <param name="name">Nom exact issu de l’inventaire des barres d’outils.</param>
        /// <returns>L’objet natif de la barre normale correspondante.</returns>
        private object FindNormalToolbar(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("ObjectName from list_toolbars is required.");
            object target = null;
            foreach (dynamic bar in vbe.CommandBars)
                if ((string)bar.Name == name)
                {
                    if (target != null) throw new InvalidOperationException("Toolbar name is ambiguous.");
                    target = (object)bar;
                }
            if (target == null) throw new InvalidOperationException("The toolbar no longer exists.");
            if ((int)((dynamic)target).Type != 0) throw new InvalidOperationException("Only normal toolbars can be changed; menu bars and popup menus are excluded.");
            return target;
        }

        /// <summary>Déplace une barre flottante en pixels ou change son ordre dans une rangée ancrée.</summary>
        /// <param name="request">Requête avec mode, coordonnées ou rangée et empreinte de disposition attendue.</param>
        /// <returns>Le résultat vérifié de l’opération et les états observés avant et après.</returns>
        public object SetToolbarPlacement(Request request)
        {
            if (request.Action != "float" && request.Action != "row")
                throw new ArgumentException("Action must be float (Left/Top pixels) or row (RowIndex).");
            bool floating = request.Action == "float";
            if (floating && (!request.ToolbarLeft.HasValue || !request.ToolbarTop.HasValue ||
                request.ToolbarLeft < -32768 || request.ToolbarLeft > 32767 || request.ToolbarTop < -32768 || request.ToolbarTop > 32767))
                throw new ArgumentException("ToolbarLeft and ToolbarTop are required integral pixels in [-32768, 32767].");
            if (!floating && (!request.RowIndex.HasValue || request.RowIndex.Value < 1))
                throw new ArgumentException("RowIndex must be a positive docking order.");
            object target = FindNormalToolbar(request.ObjectName);
            dynamic bar = target; dynamic before = ToolbarSnapshot(target);
            if (string.IsNullOrWhiteSpace(request.ExpectedToolbarLayoutVersion) || before.ToolbarLayoutVersion == null ||
                !string.Equals((string)before.ToolbarLayoutVersion, request.ExpectedToolbarLayoutVersion, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("Toolbar layout changed or could not be read. Read list_toolbars again.");
            int position = (int)before.Geometry["Position"];
            if (floating ? position != 4 : position < 0 || position > 3)
                throw new InvalidOperationException("Use set_toolbar_position first; placement does not change docking mode.");
            if (!(bool)bar.Enabled || (((int)bar.Protection & (4 | 16)) != 0))
                throw new InvalidOperationException("Toolbar disabled or protection forbids movement.");
            string error = null; object after = null; bool verified = false; bool? applied = null;
            try
            {
                if (floating) { bar.Left = request.ToolbarLeft.Value; bar.Top = request.ToolbarTop.Value; }
                else bar.RowIndex = request.RowIndex.Value;
                SaveToolbarProfile(bar, false);
            }
            catch (Exception ex) { error = ex.Message; }
            try
            {
                dynamic snapshot = ToolbarSnapshot(target); after = snapshot;
                if (snapshot.ToolbarLayoutVersion != null)
                {
                    applied = (string)snapshot.ToolbarLayoutVersion != (string)before.ToolbarLayoutVersion;
                    verified = error == null && (int)snapshot.Geometry["Position"] == position &&
                        (bool)snapshot.Properties["Visible"] == (bool)before.Properties["Visible"] &&
                        (floating ? (int)snapshot.Geometry["Left"] == request.ToolbarLeft.Value && (int)snapshot.Geometry["Top"] == request.ToolbarTop.Value :
                            (int)snapshot.Geometry["RowIndex"] == request.RowIndex.Value);
                }
            }
            catch (Exception ex) { error = error ?? ex.Message; }
            return new { Applied = applied, Verified = verified, Before = (object)before, After = after,
                NativeError = error, PersistenceVerified = false, NextRead = "list_toolbars",
                Limit = "VBE may normalize placement or rearrange neighbors. Partial changes are reported, not retried or implicitly rolled back." };
        }

        /// <summary>Ancre ou détache une barre selon l’action demandée, après contrôle de sa disposition.</summary>
        /// <param name="request">Requête avec action et empreinte de disposition attendue.</param>
        /// <returns>Le résultat de l’opération, l’état de visibilité et les instantanés observés.</returns>
        public object SetToolbarPosition(Request request)
        {
            int desired;
            switch (request.Action)
            {
                case "left": desired = 0; break;
                case "top": desired = 1; break;
                case "right": desired = 2; break;
                case "bottom": desired = 3; break;
                case "float": desired = 4; break;
                default: throw new ArgumentException("Use left, top, right, bottom or float.");
            }
            object target = FindNormalToolbar(request.ObjectName);
            dynamic selected = target;
            dynamic before = ToolbarSnapshot(target);
            if (string.IsNullOrWhiteSpace(request.ExpectedToolbarLayoutVersion) || before.ToolbarLayoutVersion == null ||
                !string.Equals((string)before.ToolbarLayoutVersion, request.ExpectedToolbarLayoutVersion, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("Toolbar layout changed or could not be read. Read list_toolbars again.");
            int original = (int)before.Geometry["Position"];
            if (original == desired) return new { Applied = false, Verified = true, Before = (object)before, After = (object)before, PersistenceVerified = false };
            if (!(bool)selected.Enabled) throw new InvalidOperationException("The toolbar is disabled.");
            int protection = (int)selected.Protection;
            if ((protection & (4 | 16)) != 0 || ((desired == 0 || desired == 2) && (protection & 32) != 0) ||
                ((desired == 1 || desired == 3) && (protection & 64) != 0))
                throw new InvalidOperationException("The toolbar protection forbids the requested movement or docking.");
            string error = null; object after = null; int? actual = null; bool? visible = null;
            try { selected.Position = desired; SaveToolbarProfile(selected, false); }
            catch (Exception ex) { error = ex.Message; }
            try { after = ToolbarSnapshot(target); actual = (int)selected.Position; visible = (bool)selected.Visible; }
            catch (Exception ex) { error = error ?? ex.Message; }
            bool visibilityPreserved = visible == (bool)before.Properties["Visible"];
            return new { Applied = actual.HasValue ? (bool?)(actual.Value != original) : null,
                Verified = error == null && actual == desired && visibilityPreserved, VisibilityPreserved = visibilityPreserved,
                Before = (object)before, After = after, NativeError = error, PersistenceVerified = false,
                NextRead = "list_toolbars", Limit = "Native docking can rearrange neighboring toolbars; exact floating coordinates and row placement are not set." };
        }
    }
}
