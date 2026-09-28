namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using CodexVBE;

    /// <summary>Contrat de scénarios établi avant exécution: options documentées FR/EN et refus de mutations non vérifiables.</summary>
    internal static class OptionsCompletionMatrix
    {
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
