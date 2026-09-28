using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Web.Script.Serialization;

namespace CodexVBE
{
    internal static partial class VbeDebugWindows
    {
        [System.Runtime.InteropServices.DllImport("user32.dll", EntryPoint="IsWindowEnabled")]
        private static extern bool NativeOptionsWindowEnabled(IntPtr window);
        /// <summary>Native enabled-state boundary, preserving the Options button validation.</summary>
        internal static Func<IntPtr, bool> OptionsWindowEnabled = NativeOptionsWindowEnabled;
        /// <summary>Sonde d’options limitée aux contrôles accessibles du dialogue courant.</summary>
        internal interface IWritableOptionsProbe : IOptionsProbe
        {
            void Write(IntPtr dialog, int tabIndex, string name, string type, object value);
            void Accept(IntPtr dialog);
        }
        /// <summary>Écrit une préférence reconnue d’édition/débogage, puis ferme par validation native.</summary>
        public static object SetVbeOption(Request request) => SetVbeOption(request, new NativeOptionsProbe());
        /// <summary>Orchestration injectable, sans modification des préférences tant que la version ne correspond pas.</summary>
        internal static object SetVbeOption(Request request, IWritableOptionsProbe native)
        {
            if (request == null || string.IsNullOrWhiteSpace(request.Pane) || string.IsNullOrWhiteSpace(request.Property) || string.IsNullOrWhiteSpace(request.ExpectedOptionsVersion))
                throw new ArgumentException("Pane, Property, Value and ExpectedOptionsVersion from read_vbe_options are required.");
            IntPtr dialog = IntPtr.Zero;
            for (int attempt = 0; attempt < 60 && dialog == IntPtr.Zero; attempt++) { native.Pause(50); dialog = native.Dialog(); }
            if (dialog == IntPtr.Zero) throw new InvalidOperationException("The native VBE Options dialog did not open.");
            bool commitRequested = false;
            try
            {
                var before = CaptureOptionsTabs(native, dialog);
                if (!string.Equals(OptionsRevision(before), request.ExpectedOptionsVersion, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("VBE options changed since inspection; read them again.");
                var names = native.Tabs(dialog);
                int index = -1;
                for (int i = 0; i < names.Count; i++) if (names[i] == request.Pane) { if (index >= 0) throw new InvalidOperationException("The options tab is ambiguous."); index = i; }
                if (index < 0) throw new InvalidOperationException("The exact options tab is absent.");
                var matches = native.Controls(dialog, index).Where(x => x.Visible && x.Enabled && x.Name == request.Property &&
                    (x.Type == "ControlType.CheckBox" || x.Type == "ControlType.RadioButton" || x.Type == "ControlType.Edit" || x.Type == "ControlType.ComboBox" || x.Type == "ControlType.List")).ToArray();
                if (matches.Length != 1 || !string.IsNullOrEmpty(matches[0].Error)) throw new InvalidOperationException("The exact option is absent, ambiguous or unreadable.");
                var selected = matches[0];
                object writeValue = ValidateEditableOption(request.Pane, selected, request.Value);
                object oldValue = selected.Value;
                native.Write(dialog, index, selected.Name, selected.Type, writeValue);
                var observed = native.Controls(dialog, index).Where(x => x.Visible && x.Enabled && x.Name == selected.Name && x.Type == selected.Type).ToArray();
                object expected = selected.Type == "ControlType.CheckBox" ? (object)((bool)writeValue ? "On" : "Off") : selected.Type == "ControlType.RadioButton" ? (object)true : Convert.ToString(writeValue, CultureInfo.InvariantCulture);
                if (observed.Length != 1 || !string.IsNullOrEmpty(observed[0].Error) || !Equals(observed[0].Value, expected))
                    throw new InvalidOperationException("The native option did not retain the requested value; Cancel will be requested.");
                commitRequested = true;
                native.Accept(dialog);
                bool closed = false;
                for (int attempt = 0; attempt < 40; attempt++) { if (native.Dialog() == IntPtr.Zero) { closed = true; break; } native.Pause(50); }
                return new { request.Pane, request.Property, Before = oldValue, After = observed[0].Value,
                    CommitRequested = true, DialogClosed = closed, ControlValueVerified = true,
                    PersistenceVerified = false, NextRead = "read_vbe_options",
                    Limit = "Reopen Options to verify committed preferences. A requested OK or a closed dialog alone is not a restart persistence proof. Do not retry automatically." };
            }
            finally
            {
                if (!commitRequested || native.Dialog() != IntPtr.Zero) native.Close(dialog);
            }
        }
        /// <summary>Valide une préférence reconnue; les listes exigent un choix natif exact et unique.</summary>
        internal static object ValidateEditableOption(string tab, OptionsControl control, object value)
        {
            string normalized = (control.Name ?? "").Replace("&", "").Trim().TrimEnd(':').Trim().ToLowerInvariant();
            string tabName = (tab ?? "").Replace("&", "").Trim().ToLowerInvariant();
            bool editor = tabName == "editor" || tabName == "éditeur" || tabName == "editeur";
            bool general = tabName == "general" || tabName == "général";
            bool format = new[] { "editor format", "format de l'éditeur", "format de l’éditeur", "format de l'editeur" }.Contains(tabName);
            bool docking = tabName == "docking" || tabName == "ancrage";
            if (!editor && !general && !format && !docking) throw new InvalidOperationException("The native options tab is not supported.");
            var editorChecks = new[] { "auto syntax check", "require variable declaration", "auto list members", "auto quick info", "auto data tips", "auto indent",
                "drag-and-drop text editing", "default to full module view", "procedure separator",
                "vérification automatique de la syntaxe", "déclaration des variables obligatoire", "exiger une déclaration de variable", "liste des membres automatique", "liste automatique des membres", "info rapide automatique", "informations rapides automatiques", "complément automatique des instructions", "info express automatique", "info-bulles automatiques", "conseils sur les données automatiques", "retrait automatique",
                "modification de texte par glisser-déplacer", "édition de texte par glisser-déplacer", "affichage module complet par défaut", "affichage du module complet par défaut", "séparateur de procédure", "séparation des procédures" };
            var generalChecks = new[] { "compile on demand", "background compile", "show grid", "align controls to grid", "notify before state loss", "show tooltips", "collapse proj. hides windows",
                "compilation à la demande", "compilation sur demande", "compiler à la demande", "compilation en arrière-plan", "compiler en arrière-plan", "afficher la grille", "aligner les contrôles sur la grille", "avertir avant la perte d'état", "avertir avant la perte d’état", "notifier avant la perte d’état", "notifier avant la perte d'état", "afficher les info-bulles", "réduire le proj. masque les fenêtres", "réduire le projet masque les fenêtres" };
            var formatChecks = new[] { "margin indicator bar", "barre des indicateurs en marge", "barre d'indicateurs en marge", "barre d’indicateurs en marge" };
            var dockChecks = new[] { "immediate window", "locals window", "watch window", "project explorer", "properties window", "object browser", "toolbox",
                "fenêtre exécution", "fenêtre variables locales", "fenêtre espions", "explorateur de projets", "explorateur de projet", "fenêtre propriétés", "explorateur d'objets", "explorateur d’objets", "boîte à outils" };
            if (control.Type == "ControlType.CheckBox" && (editor ? editorChecks : general ? generalChecks : format ? formatChecks : dockChecks).Contains(normalized))
            { if (!(value is bool)) throw new ArgumentException("This option requires a boolean Value."); return value; }
            if (general && control.Type == "ControlType.RadioButton" &&
                new[] { "break on all errors", "break in class module", "break on unhandled errors", "arrêt sur toutes les erreurs", "arrêt dans le module de classe", "arrêt sur les erreurs non gérées" }.Contains(normalized))
            { if (!(value is bool) || !(bool)value) throw new ArgumentException("Select an error-trapping radio option with Value=true."); return true; }
            bool width = editor && new[] { "tab width", "largeur de tabulation", "largeur de la tabulation" }.Contains(normalized);
            bool grid = general && new[] { "width", "height", "largeur", "hauteur" }.Contains(normalized);
            if (control.Type == "ControlType.Edit" && (width || grid))
            {
                string text = Convert.ToString(value, CultureInfo.InvariantCulture);
                int minimum = grid ? 2 : 1, maximum = grid ? 60 : 32;
                if (!int.TryParse(text, NumberStyles.None, CultureInfo.InvariantCulture, out int size) || size < minimum || size > maximum)
                    throw new ArgumentException("The option must be an integer between " + minimum + " and " + maximum + ".");
                return size.ToString(CultureInfo.InvariantCulture);
            }
            bool formatChoice = format && ((control.Type == "ControlType.ComboBox" &&
                new[] { "font", "police", "size", "taille", "foreground", "premier plan", "background", "arrière-plan", "indicator", "indicateur" }.Contains(normalized)) ||
                (control.Type == "ControlType.List" && new[] { "code colors", "couleurs du code", "color text", "texte couleur" }.Contains(normalized)));
            if (formatChoice)
            {
                if (!(value is string choice) || string.IsNullOrWhiteSpace(choice)) throw new ArgumentException("Select an exact string from the observed native Choices.");
                if (control.Choices == null || control.Choices.Count(x => string.Equals(x, choice, StringComparison.Ordinal)) != 1)
                    throw new InvalidOperationException("The exact native choice is absent, ambiguous or unreadable.");
                return choice;
            }
            throw new InvalidOperationException("This native option is not in the supported preference list.");
        }
        /// <summary>Capture bornée commune à la lecture et au contrôle de version avant écriture.</summary>
        private static List<object> CaptureOptionsTabs(IOptionsProbe native, IntPtr dialog)
        {
            var tabs = new List<object>(); var names = native.Tabs(dialog);
            if (names.Count < 1 || names.Count > 8) throw new InvalidOperationException("Unexpected native VBE Options tab count: " + names.Count + ".");
            for (int i = 0; i < names.Count; i++)
            {
                if (string.IsNullOrWhiteSpace(names[i])) throw new InvalidOperationException("A native VBE Options tab is unreadable.");
                var observed = native.Controls(dialog, i);
                if (observed.Count > 2000) throw new InvalidOperationException("The Options dialog has too many controls to inspect safely: " + observed.Count + ".");
                var controls = observed.Where(x => x.Visible && x.Enabled && (!string.IsNullOrWhiteSpace(x.Name) || x.Type != "ControlType.Text"))
                    .Select(x => (object)new { x.Name, x.Type, x.Value, x.Error, x.Choices }).ToList();
                tabs.Add(new { Tab = names[i], Controls = controls, Count = controls.Count });
            }
            return tabs;
        }
        /// <summary>Version du contenu visible des options, indépendante des handles transitoires.</summary>
        private static string OptionsRevision(List<object> tabs)
        {
            using (var sha = SHA256.Create()) return BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(new JavaScriptSerializer().Serialize(tabs)))).Replace("-", "").ToLowerInvariant();
        }
    }
}
