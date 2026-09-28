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
                var matches = native.Controls(dialog, index).Where(x => x.Visible && x.Enabled && x.Name == request.Property).ToArray();
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
        /// <summary>Autorise seulement les préférences connues d’éditeur et de débogueur, sans format de couleurs ni sécurité.</summary>
        internal static object ValidateEditableOption(string tab, OptionsControl control, object value)
        {
            string normalized = (control.Name ?? "").Replace("&", "").Trim().TrimEnd(':').ToLowerInvariant();
            string tabName = (tab ?? "").Replace("&", "").Trim().ToLowerInvariant();
            bool editor = tabName == "editor" || tabName == "éditeur" || tabName == "editeur";
            bool general = tabName == "general" || tabName == "général";
            if (!editor && !general) throw new InvalidOperationException("Only recognized Editor/General options are writable; formatting, docking and security are excluded.");
            var editorChecks = new[] { "auto syntax check", "require variable declaration", "auto list members", "auto quick info", "auto data tips", "auto indent",
                "vérification automatique de la syntaxe", "déclaration des variables obligatoire", "liste des membres automatique", "info rapide automatique", "info-bulles automatiques", "retrait automatique" };
            var generalChecks = new[] { "compile on demand", "background compile", "compilation à la demande", "compilation en arrière-plan", "compiler en arrière-plan" };
            if (control.Type == "ControlType.CheckBox" && (editor ? editorChecks : generalChecks).Contains(normalized))
            { if (!(value is bool)) throw new ArgumentException("This option requires a boolean Value."); return value; }
            if (general && control.Type == "ControlType.RadioButton" &&
                new[] { "break on all errors", "break in class module", "break on unhandled errors", "arrêt sur toutes les erreurs", "arrêt dans le module de classe", "arrêt sur les erreurs non gérées" }.Contains(normalized))
            { if (!(value is bool) || !(bool)value) throw new ArgumentException("Select an error-trapping radio option with Value=true."); return true; }
            if (editor && control.Type == "ControlType.Edit" && new[] { "tab width", "largeur de tabulation" }.Contains(normalized))
            {
                string text = Convert.ToString(value, CultureInfo.InvariantCulture);
                if (!int.TryParse(text, NumberStyles.None, CultureInfo.InvariantCulture, out int size) || size < 1 || size > 32)
                    throw new ArgumentException("Tab width must be an integer between 1 and 32.");
                return size.ToString(CultureInfo.InvariantCulture);
            }
            throw new InvalidOperationException("This native option is not in the supported editing/debugging preference list.");
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
                    .Select(x => (object)new { x.Name, x.Type, x.Value, x.Error }).ToList();
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
