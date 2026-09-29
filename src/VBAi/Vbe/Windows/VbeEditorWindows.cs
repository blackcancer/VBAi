using System;
using System.Collections.Generic;

namespace VBAi
{
    // Reads the VBIDE collections directly. In particular, do not use
    // VBComponent.CodePane for inspection: its getter opens and activates a pane.
    /// <summary>Crée le lecteur pour l’instance VBE fournie.</summary>
    internal sealed partial class VbeEditorWindows
    {
        /// <summary>Instance VBE dont les collections sont inspectées.</summary>
        private readonly dynamic vbe;

        /// <summary>Crée le lecteur pour l’instance VBE fournie.</summary>
        /// <param name="vbe">Instance VBE dont les collections seront lues.</param>
        public VbeEditorWindows(object vbe) { this.vbe = vbe; }

        /// <summary>Retourne les fenêtres VBE et la fenêtre active, en isolant les erreurs de lecture.</summary>
        /// <returns>Liste des fenêtres avec l’état de la fenêtre active et les erreurs de lecture.</returns>
        public object Windows()
        {
            var windows = new List<object>();
            int index = 0;
            foreach (dynamic window in vbe.Windows)
                windows.Add(WindowSnapshot(window, ++index));
            object active = null;
            try
            {
                dynamic window = vbe.ActiveWindow;
                if (window != null) active = WindowSnapshot(window, null);
            }
            catch (Exception ex) { active = new { Error = ex.Message }; }
            return new { Windows = windows, ActiveWindow = active };
        }

        /// <summary>Retourne les volets de code VBE et le volet actif sans en ouvrir de nouveau.</summary>
        /// <returns>Liste des volets de code et du volet actif avec leurs erreurs de lecture.</returns>
        public object CodePanes()
        {
            var panes = new List<object>();
            int index = 0;
            foreach (dynamic pane in vbe.CodePanes)
                panes.Add(PaneSnapshot(pane, ++index));
            object active = null;
            try
            {
                dynamic pane = vbe.ActiveCodePane;
                if (pane != null) active = PaneSnapshot(pane, null);
            }
            catch (Exception ex) { active = new { Error = ex.Message }; }
            return new { CodePanes = panes, ActiveCodePane = active };
        }

        /// <summary>Lit les propriétés principales de l’instance et collecte les erreurs par propriété.</summary>
        /// <returns>Propriétés d’environnement accessibles et erreurs par propriété.</returns>
        public object Environment()
        {
            var fields = new Dictionary<string, object>();
            var errors = new Dictionary<string, string>();
            Read(fields, errors, "Version", () => (string)vbe.Version);
            Read(fields, errors, "ProjectCount", () => (int)vbe.VBProjects.Count);
            Read(fields, errors, "WindowCount", () => (int)vbe.Windows.Count);
            Read(fields, errors, "CodePaneCount", () => (int)vbe.CodePanes.Count);
            Read(fields, errors, "AddInCount", () => (int)vbe.AddIns.Count);
            Read(fields, errors, "ActiveProject", () => (string)vbe.ActiveVBProject.Name);
            return new { Properties = fields, Errors = errors };
        }

        /// <summary>Liste les add-ins enregistrés dans la collection VBE.AddIns.</summary>
        /// <returns>Add-ins VBE, nombre total et périmètre de la collection inspectée.</returns>
        public object AddIns()
        {
            var addIns = new List<object>();
            int index = 0;
            foreach (dynamic addIn in vbe.AddIns)
                addIns.Add(AddInSnapshot(addIn, ++index));
            return new { AddIns = addIns, Count = addIns.Count,
                Scope = "VBE.AddIns contains VBE-registered add-ins, not the host application's COMAddIns." };
        }

        /// <summary>Active une fenêtre exacte déjà visible et rapporte le résultat de vérification.</summary>
        /// <param name="caption">Légende exacte de la fenêtre telle que renvoyée par vbe_windows.</param>
        /// <param name="type">Type exact de la fenêtre telle que renvoyée par vbe_windows.</param>
        /// <returns>Résultat d’activation et état de vérification de la fenêtre active.</returns>
        public object FocusWindow(string caption, int type)
        {
            dynamic target = FindExactWindow(caption, type);
            if (!(bool)target.Visible)
                throw new InvalidOperationException("The requested window is hidden; SetFocus requires a visible window.");
            target.SetFocus();
            dynamic active = vbe.ActiveWindow;
            bool verified = active != null &&
                string.Equals((string)active.Caption, caption, StringComparison.Ordinal) &&
                (int)active.Type == type;
            return new { WindowCaption = caption, WindowType = type, SetFocusInvoked = true,
                Verification = verified ? "ActiveWindowReadback" : "Unverified",
                ActiveWindow = active == null ? null : WindowSnapshot(active, null) };
        }

        /// <summary>Rend visible puis active une fenêtre VBE exacte.</summary>
        /// <param name="caption">Légende exacte de la fenêtre telle que renvoyée par vbe_windows.</param>
        /// <param name="type">Type exact de la fenêtre telle que renvoyée par vbe_windows.</param>
        /// <returns>Visibilité avant/après et état de vérification de la fenêtre active.</returns>
        public object ShowWindow(string caption, int type)
        {
            dynamic target = FindExactWindow(caption, type);
            bool wasVisible = (bool)target.Visible;
            if (!wasVisible) target.Visible = true;
            bool nowVisible = (bool)target.Visible;
            if (!nowVisible)
                throw new InvalidOperationException("The native VBE window did not become visible.");
            target.SetFocus();
            dynamic active = vbe.ActiveWindow;
            bool activeVerified = active != null &&
                string.Equals((string)active.Caption, caption, StringComparison.Ordinal) &&
                (int)active.Type == type;
            return new { WindowCaption = caption, WindowType = type, WasVisible = wasVisible,
                Visible = nowVisible, FocusVerified = activeVerified,
                ActiveWindow = active == null ? null : WindowSnapshot(active, null) };
        }

        /// <summary>Retourne l’état de visibilité et les fenêtres liées au cadre de la cible.</summary>
        /// <param name="caption">Légende exacte de la fenêtre telle que renvoyée par vbe_windows.</param>
        /// <param name="type">Type exact de la fenêtre telle que renvoyée par vbe_windows.</param>
        /// <returns>Visibilité, cadre parent et fenêtres liées, avec les erreurs rencontrées.</returns>
        public object WindowLinkage(string caption, int type)
        {
            dynamic target = FindExactWindow(caption, type);
            var fields = new Dictionary<string, object>();
            var errors = new Dictionary<string, string>();
            Read(fields, errors, "Visible", () => (bool)target.Visible);
            try
            {
                dynamic frame = target.LinkedWindowFrame;
                fields["IsLinked"] = frame != null;
                if (frame != null)
                {
                    Read(fields, errors, "FrameCaption", () => (string)frame.Caption);
                    var linked = new List<object>();
                    foreach (dynamic window in frame.LinkedWindows)
                        linked.Add(new { Caption = (string)window.Caption, Type = (int)window.Type });
                    fields["LinkedWindows"] = linked;
                }
            }
            catch (Exception ex) { errors["LinkedWindowFrame"] = ex.Message; }
            return new { WindowCaption = caption, WindowType = type,
                Properties = fields, Errors = errors };
        }

        /// <summary>Ferme une fenêtre VBE exacte et vérifie si elle a disparu ou est cachée.</summary>
        /// <param name="caption">Légende exacte de la fenêtre telle que renvoyée par vbe_windows.</param>
        /// <param name="type">Type exact de la fenêtre telle que renvoyée par vbe_windows.</param>
        /// <returns>État de fermeture vérifié et occurrences encore présentes.</returns>
        public object CloseWindow(string caption, int type)
        {
            if (string.Equals(caption, "VBAi", StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The VBAi tool window cannot close itself through this command.");
            dynamic target = FindExactWindow(caption, type);
            if (!(bool)target.Visible)
                throw new InvalidOperationException("The requested VBE window is already hidden.");
            target.Close();
            int matches = 0;
            bool visible = false;
            foreach (dynamic window in vbe.Windows)
                if (string.Equals((string)window.Caption, caption, StringComparison.Ordinal) &&
                    (int)window.Type == type)
                { matches++; visible |= (bool)window.Visible; }
            return new { WindowCaption = caption, WindowType = type, CloseInvoked = true,
                Verification = matches == 0 ? "RemovedFromWindows" :
                    matches == 1 && !visible ? "HiddenInWindows" : "Unverified",
                RemainingMatches = matches, RemainingVisible = visible };
        }

        /// <summary>Résout une fenêtre selon sa légende et son type, en refusant les doublons.</summary>
        /// <param name="caption">Légende exacte de la fenêtre telle que renvoyée par vbe_windows.</param>
        /// <param name="type">Type exact de la fenêtre telle que renvoyée par vbe_windows.</param>
        /// <returns>Objet fenêtre qui correspond à la légende et au type uniques.</returns>
        private dynamic FindExactWindow(string caption, int type)
        {
            if (string.IsNullOrWhiteSpace(caption) || type < 0)
                throw new ArgumentException("WindowCaption and WindowType from vbe_windows are required.");
            dynamic target = null;
            foreach (dynamic window in vbe.Windows)
            {
                if (!string.Equals((string)window.Caption, caption, StringComparison.Ordinal) ||
                    (int)window.Type != type) continue;
                if (target != null)
                    throw new InvalidOperationException("More than one VBE window matches the requested caption and type.");
                target = window;
            }
            if (target == null) throw new InvalidOperationException("The requested VBE window is no longer present.");
            return target;
        }

        /// <summary>Lit les propriétés d’une fenêtre, en enregistrant séparément les erreurs de COM.</summary>
        /// <param name="window">Fenêtre COM à inspecter.</param>
        /// <param name="index">Index de la fenêtre dans la collection, ou null pour la fenêtre active.</param>
        /// <returns>Dictionnaires des propriétés lues et des erreurs COM rencontrées.</returns>
        private static object WindowSnapshot(dynamic window, int? index)
        {
            var fields = new Dictionary<string, object>();
            var errors = new Dictionary<string, string>();
            Read(fields, errors, "Caption", () => (string)window.Caption);
            Read(fields, errors, "Type", () => (int)window.Type);
            Read(fields, errors, "Visible", () => (bool)window.Visible);
            Read(fields, errors, "WindowState", () => (int)window.WindowState);
            Read(fields, errors, "Left", () => (int)window.Left);
            Read(fields, errors, "Top", () => (int)window.Top);
            Read(fields, errors, "Width", () => (int)window.Width);
            Read(fields, errors, "Height", () => (int)window.Height);
            return new { Index = index, Properties = fields, Errors = errors };
        }

        /// <summary>Lit les informations du volet de code et sa sélection courante.</summary>
        /// <param name="pane">Volet de code COM à inspecter.</param>
        /// <param name="index">Index de la fenêtre dans la collection, ou null pour la fenêtre active.</param>
        /// <returns>Dictionnaires des propriétés lues, de la sélection et des erreurs COM.</returns>
        private static object PaneSnapshot(dynamic pane, int? index)
        {
            var fields = new Dictionary<string, object>();
            var errors = new Dictionary<string, string>();
            Read(fields, errors, "Module", () => (string)pane.CodeModule.Parent.Name);
            Read(fields, errors, "Project", () => (string)pane.CodeModule.Parent.Collection.Parent.Name);
            Read(fields, errors, "ProjectPath", () => (string)pane.CodeModule.Parent.Collection.Parent.FileName);
            Read(fields, errors, "CodePaneView", () => (int)pane.CodePaneView);
            Read(fields, errors, "TopLine", () => (int)pane.TopLine);
            Read(fields, errors, "CountOfVisibleLines", () => (int)pane.CountOfVisibleLines);
            Read(fields, errors, "WindowCaption", () => (string)pane.Window.Caption);
            try
            {
                int startLine = 0, startColumn = 0, endLine = 0, endColumn = 0;
                pane.GetSelection(ref startLine, ref startColumn, ref endLine, ref endColumn);
                fields["Selection"] = new { StartLine = startLine, StartColumn = startColumn,
                    EndLine = endLine, EndColumn = endColumn };
            }
            catch (Exception ex) { errors["Selection"] = ex.Message; }
            return new { Index = index, Properties = fields, Errors = errors };
        }

        /// <param name="fields">Dictionnaire des valeurs lues avec succès.</param>
        /// <param name="errors">Dictionnaire recevant les erreurs indexées par propriété.</param>
        /// <param name="name">Nom de la propriété à lire.</param>
        /// <param name="getter">Accès COM à exécuter pour lire la valeur.</param>
        /// <summary>Lit une propriété COM et conserve séparément l’exception éventuelle.</summary>
        private static void Read(IDictionary<string, object> fields,
            IDictionary<string, string> errors, string name, Func<object> getter)
        {
            try { fields[name] = getter(); }
            catch (Exception ex) { errors[name] = ex.Message; }
        }
    }
}
