using System;
using System.Collections.Generic;
using System.IO;
using System.Globalization;
using System.Runtime.CompilerServices;
using System.Windows.Forms;
#if VBAI_UPDATER
using UiText = VBAi.UpdateText;
#endif

namespace VBAi
{
    /// <summary>Opens the localized local manual and routes ordinary form help requests to the relevant chapter.</summary>
    internal static class UiHelp
    {
        /// <summary>Marks forms already subscribed; keys do not retain disposed forms.</summary>
        private static readonly ConditionalWeakTable<Form, object> Attached = new ConditionalWeakTable<Form, object>();

        /// <summary>Topics for top-level product interfaces; unknown types use the manual entry page.</summary>
        private static readonly Dictionary<string, string> Topics = new Dictionary<string, string> {
            { "ChatWindow", "conversation" }, { "LlmSettingsWindow", "settings" },
            { "GitWindow", "git-conversion" }, { "ModernEditorWindow", "editor" },
            { "ProjectAccessWindow", "privacy" }, { "VbeApprovalDialog", "privacy" },
            { "TestExplorerWindow", "tests" }, { "TestSupportReviewDialog", "tests" },
            { "UpdateWindow", "updates" }, { "UpdateProgressWindow", "updates" },
            { "CrashReportWindow", "support" }, { "AboutWindow", "start" }
        };

        /// <summary>Invokes the Windows HTML Help viewer; replaceable for offline interaction tests.</summary>
        internal static Action<IWin32Window, string, string> Launch = (owner, file, topic) =>
            Help.ShowHelp(owner as Control, file, HelpNavigator.Topic, topic + ".html");

        /// <summary>Checks packaged paths only when help is explicitly requested.</summary>
        internal static Func<string, bool> Exists = File.Exists;

        /// <summary>Explains an absent local manual without navigating to an external site automatically.</summary>
        internal static Action<IWin32Window> ShowUnavailable = owner => MessageBox.Show(owner,
            UiText.Get("Local help is not included in this build. Place the VBAi help files in the Help folder beside the application."),
            "VBAi", MessageBoxButtons.OK, MessageBoxIcon.Information);

        /// <summary>Resolves the packaged manual beside the add-in assembly, independent of the host working directory.</summary>
        /// <value>Absolute path to the current supported interface language in the application's Help folder.</value>
        internal static string FilePath => PathFor(InterfaceCulture);

        /// <summary>Uses the language selected by the VBE, or the independent updater's culture.</summary>
        private static CultureInfo InterfaceCulture
        {
            get
            {
#if VBAI_UPDATER
                return CultureInfo.GetCultureInfo(UpdateText.Culture);
#else
                return UiText.Culture;
#endif
            }
        }

        /// <summary>Maps a culture to a supported, fixed local filename; it cannot supply arbitrary paths.</summary>
        /// <param name="culture">Requested UI culture, including regional variants.</param>
        /// <returns>Absolute packaged path for its supported language.</returns>
        internal static string PathFor(CultureInfo culture) => Path.Combine(
            Path.GetDirectoryName(typeof(UiHelp).Assembly.Location), "Help",
            "VBAi." + UiLanguages.For(culture).CultureName + ".chm");

        /// <summary>Selects installed localized help with English, then French, as explicit offline fallbacks.</summary>
        /// <param name="culture">Current interface culture.</param>
        /// <returns>The available local archive path, or null if none exists.</returns>
        internal static string Resolve(CultureInfo culture)
        {
            var paths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var candidate in new[] { PathFor(culture), PathFor(CultureInfo.GetCultureInfo("en-US")), PathFor(CultureInfo.GetCultureInfo("fr-FR")) })
                if (paths.Add(candidate) && Exists(candidate)) return candidate;
            return null;
        }

        /// <summary>Selects a chapter without inspecting project data or invoking a live service.</summary>
        /// <param name="formType">Form type whose help should open; null selects the entry page.</param>
        /// <returns>A fixed, local HTML topic identifier.</returns>
        internal static string TopicFor(Type formType) => formType != null && Topics.TryGetValue(formType.Name, out var topic) ? topic : "start";

        /// <summary>Attaches one help event to a live form; Designer construction remains inert.</summary>
        /// <param name="form">Form on its owning UI thread.</param>
        internal static void Attach(Form form)
        {
            if (form == null || UiTheme.IsDesignPreview(form) || Attached.TryGetValue(form, out _)) return;
            Attached.Add(form, new object());
            form.HelpRequested += (sender, args) => { args.Handled = true; Open(form, TopicFor(form.GetType())); };
        }

        /// <summary>Opens a fixed manual chapter, or tells the user how to obtain the local help when it is absent.</summary>
        /// <param name="owner">Native owner for the help request and any missing-help notice.</param>
        /// <param name="topic">Allowlisted topic, or start; arbitrary paths and URLs are never forwarded.</param>
        internal static void Open(IWin32Window owner, string topic = "start")
        {
            if (topic != "start" && !Topics.ContainsValue(topic)) topic = "start";
            var file = Resolve(InterfaceCulture);
            if (file == null)
            {
                ShowUnavailable(owner);
                return;
            }
            Launch(owner, file, topic);
        }
    }
}
