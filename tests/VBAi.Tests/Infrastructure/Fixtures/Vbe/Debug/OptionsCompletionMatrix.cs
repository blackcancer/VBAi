namespace VBAi.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using VBAi;

    /// <summary>Contrat de scénarios établi avant exécution: options documentées FR/EN et refus de mutations non vérifiables.</summary>
    internal static class OptionsCompletionMatrix
    {
        /// <summary>Scénarios de listes natives établis avant mutations : police, taille vide, palettes opaques et gardes.</summary>
        internal static readonly string[] NativeFormatScenarios = {
            "font exact edit value with no selected index", "font duplicate labels refuse write",
            "size empty catalogue readable but refuses write", "dropdown populated and closed",
            "palette seventeen real indices with unlabeled entries", "opaque selected index readback",
            "foreign PID rejected", "non ComboBox rejected", "owner data without strings rejected",
            "negative/oversized counts rejected", "oversized strings rejected", "invalid selected index rejected",
            "native selection failure cancels", "palette index revision invalidates write",
            "native labels revision invalidates write", "full options version restored after native qualification"
        };
        /// <summary>Optimisation : lecture palettes seules par catégorie, dernière sélection restaurée et erreurs toujours propagées.</summary>
        internal static readonly string[] CategoryInspectionScenarios = { "ten palette-only reads", "disabled palette preserved",
            "original category restored", "palette read throws", "selection throws", "null palette", "wrong palette count",
            "wrong palette type", "unknown palette name", "duplicate palette", "unreadable palette", "missing value", "offscreen palette",
            "empty/duplicate/unreadable catalogue refuses before selection", "original selection outside catalogue refuses" };
        /// <summary>Notification catégorie : parent exact, LBN_SELCHANGE, identité invalidée et erreur propagée sans second envoi.</summary>
        internal static readonly string[] CategoryNotificationScenarios = { "owned parent dispatch", "zero list", "zero parent", "same handles",
            "negative ID", "oversized ID", "missing dispatcher", "dispatch failure propagates once" };
        /// <summary>Fixtures UIA : vrais HWND/PID/classe dialogue, garde onglet et maintien des refus pattern/valeur/collection.</summary>
        internal static readonly string[] OwnedOptionsFixtureScenarios = { "real owned #32770 dialog", "UIA root HWND matches dialog",
            "tab provider own PID", "exact native choices", "no typing/focus", "missing/read-only/password pattern refuses",
            "unavailable element skipped", "2001 control collection refuses", "direct CreateWindowEx preserves standard class without WinForms prefix",
            "UIA NativeWindow subclass preserves native identity", "raw child HWNDs remain inside their actual parent dialog" };
        /// <summary>Contrôles Win32 réels détenus : Combo lecture/écriture, notifications ListBox, catégories et fautes de transport injectées.</summary>
        internal static readonly string[] OwnedNativeOptionsScenarios = {
            "real ComboBox text and exact selection", "real empty editable ComboBox and dropdown restoration",
            "unlabeled real entries expose native indices", "duplicate/missing text refuses write", "zero/wrong-class/foreign-PID guards",
            "HASSTRINGS guard", "count negative/oversized", "label negative/oversized/changed", "selected index invalid",
            "edit value negative/oversized/changed", "native selection fails before notification", "owned parent SELCHANGE/SELENDOK",
            "real ListBox category notification updates owner palettes", "all categories capture and restore", "disabled palette state retained",
            "absent/duplicate list", "wrong/unreadable/ambiguous category selection", "missing pattern/empty/duplicate/oversized categories",
            "foreign category provider PID", "palette missing/renamed/offscreen/password/wrong kind/handle changed",
            "palette failure restores original category", "tab index/PID/pattern/retention guards"
        };
        /// <summary>Cas positifs: chaque préférence ajoutée doit passer version, écriture, relecture et validation.</summary>
        internal static IEnumerable<Tuple<string, string, string, object>> Supported()
        {
            foreach (string name in new[] { "Drag-and-Drop Text Editing", "Default to Full Module View", "Procedure Separator" })
                yield return Row("Editor", name, "ControlType.CheckBox", true);
            foreach (string name in new[] { "Modification de texte par glisser-déplacer", "Affichage module complet par défaut", "Séparateur de procédure" })
                yield return Row("Éditeur", name, "ControlType.CheckBox", true);
            foreach (string name in new[] { "Show Grid", "Align Controls to Grid", "Notify Before State Loss", "Show ToolTips", "Collapse Proj. Hides Windows" })
                yield return Row("General", name, "ControlType.CheckBox", false);
            foreach (string name in new[] { "Afficher la grille", "Aligner les contrôles sur la grille", "Notifier avant la perte d’état", "Afficher les info-bulles", "Réduire le proj. masque les fenêtres" })
                yield return Row("Général", name, "ControlType.CheckBox", false);
            foreach (string name in new[] { "Width", "Height" })
                foreach (int bound in new[] { 2, 60 }) yield return Row("General", name, "ControlType.Edit", bound);
            foreach (string name in new[] { "Largeur", "Hauteur" })
                foreach (int bound in new[] { 2, 60 }) yield return Row("Général", name, "ControlType.Edit", bound);
            yield return Row("Editor Format", "Margin Indicator Bar", "ControlType.CheckBox", false);
            yield return Row("Format de l’éditeur", "Barre des indicateurs en marge", "ControlType.CheckBox", false);
            foreach (string name in new[] { "Font", "Size", "Foreground", "Background", "Indicator" })
                yield return Row("Editor Format", name, "ControlType.ComboBox", "Native choice");
            foreach (string name in new[] { "Police", "Taille", "Premier plan", "Arrière-plan", "Indicateur" })
                yield return Row("Format de l'éditeur", name, "ControlType.ComboBox", "Native choice");
            yield return Row("Editor Format", "Code Colors", "ControlType.List", "Native choice");
            yield return Row("Format de l’éditeur", "Couleurs du code", "ControlType.List", "Native choice");
            foreach (string name in new[] { "Immediate Window", "Locals Window", "Watch Window", "Project Explorer", "Properties Window", "Object Browser", "Toolbox" })
                yield return Row("Docking", name, "ControlType.CheckBox", true);
            foreach (string name in new[] { "Fenêtre Exécution", "Fenêtre Variables locales", "Fenêtre Espions", "Explorateur de projets", "Fenêtre Propriétés", "Explorateur d’objets", "Boîte à outils" })
                yield return Row("Ancrage", name, "ControlType.CheckBox", true);
        }

        /// <summary>Refus avant écriture pour borne, format numérique ou valeur absente non documentés.</summary>
        internal static readonly object[] InvalidGridValues = { null, 1, 61, -2, 2.5, true, "2.0", " 2", "2\n" };

        /// <summary>Refus avant écriture: absent, doublon, catalogue absent ou illisible, mauvais type, onglet incorrect, nom inconnu.</summary>
        internal static readonly string[] InvalidChoiceScenarios = { "absent", "duplicate", "null catalogue", "empty catalogue", "null value", "numeric value", "wrong tab", "unknown name", "wrong type", "unreadable" };

        /// <summary>Mutations concurrentes: choix, catégorie ou sélection changent après inspection; aucune écriture ni validation.</summary>
        internal static readonly string[] RevisionScenarios = { "choices", "selection", "category" };

        private static Tuple<string, string, string, object> Row(string tab, string name, string type, object value)
        {
            return Tuple.Create(tab, name, type, value);
        }
    }
}
