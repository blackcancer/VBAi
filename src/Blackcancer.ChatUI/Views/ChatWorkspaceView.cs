// SPDX-License-Identifier: MPL-2.0
// Fixed Designer-backed workspace ported from SolidAi's native ChatPaneControl.
using System;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Windows.Forms;

namespace Blackcancer.ChatUI.Views
{
    /// <summary>The common workspace surface; hosts bind domain actions to its public parts.</summary>
    [ToolboxItem(true)]
    public sealed partial class ChatWorkspaceView : UserControl
    {
        public ChatWorkspaceView() : this(CultureInfo.CurrentUICulture) { }

        public ChatWorkspaceView(CultureInfo culture)
        {
            UiCulture = culture ?? throw new ArgumentNullException(nameof(culture));
            LoadResources(InitializeComponent);
            if (!UiCulture.Equals(CultureInfo.CurrentUICulture)) ApplyCultureResources();
            DoubleBuffered = true;
        }
        private void ApplyCultureResources()
        {
            LoadResources(ApplyCultureResourcesCore);
        }

        /// <summary>Loads a workspace's resources without requiring redirects in the host application's configuration.</summary>
        private static void LoadResources(Action load)
        {
            const string currentReader = "System.Resources.Extensions, Version=8.0.0.0, Culture=neutral, PublicKeyToken=cc7b13ffcd2ddd51";
            const string legacyReader = "System.Resources.Extensions, Version=4.0.0.0, Culture=neutral, PublicKeyToken=cc7b13ffcd2ddd51";
            // COM hosts need not probe the add-in directory. Load its exact reader before
            // registering the handler; a typeof(reader) inside that handler can itself
            // require an unresolved reader assembly before its name filter executes.
            string directory = Path.GetDirectoryName(typeof(ChatWorkspaceView).Assembly.Location);
            if (string.IsNullOrEmpty(directory)) throw new FileNotFoundException("The shared chat assembly has no resource-reader directory.");
            string path = Path.Combine(directory, "System.Resources.Extensions.dll");
            if (!string.Equals(AssemblyName.GetAssemblyName(path).FullName, currentReader, StringComparison.Ordinal))
                throw new FileLoadException("The shared chat resource reader has an unexpected assembly identity.", path);
            Assembly resourceReader = Assembly.LoadFrom(path);
            if (!string.Equals(resourceReader.GetName().FullName, currentReader, StringComparison.Ordinal))
                throw new FileLoadException("The loaded shared chat resource reader has an unexpected assembly identity.", path);
            ResolveEventHandler reader = (sender, request) =>
                request.Name == legacyReader || request.Name == currentReader ? resourceReader : null;
            AppDomain.CurrentDomain.AssemblyResolve += reader;
            try { load(); }
            finally { AppDomain.CurrentDomain.AssemblyResolve -= reader; }
        }

        private void ApplyCultureResourcesCore()
        {
            // The Designer uses the standard two-argument ApplyResources calls.
            // Culture selection belongs to the host adapter at runtime, per instance.
            var resources = new ComponentResourceManager(typeof(ChatWorkspaceView));
            foreach (var field in GetType().GetFields(System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance))
            {
                object value = field.GetValue(this);
                if (value != null) resources.ApplyResources(value, field.Name, UiCulture);
            }
            resources.ApplyResources(this, "$this", UiCulture);
            // Restore left-to-right layout after leaving an RTL culture.
            RightToLeft = UiCulture.TextInfo.IsRightToLeft ? RightToLeft.Yes : RightToLeft.No;
        }

        [Browsable(false), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public CultureInfo UiCulture { get; private set; }
        public void ApplyCulture(CultureInfo culture)
        {
            if (culture == null) throw new ArgumentNullException(nameof(culture));
            if (UiCulture.Equals(culture)) return;
            UiCulture = culture;
            ApplyCultureResources();
        }

        [Browsable(false), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public ChatWorkspaceView Parts => this;
        /// <summary>Expose an optional host extension without relying on positional layout indices.</summary>
        public void SetExtensionVisible(Control extension, bool visible)
        {
            if (extension != AdditionalContextHost && extension != PendingMessagesHost)
                throw new ArgumentException("The control must be a workspace extension slot.", nameof(extension));
            int row = rootLayout.GetRow(extension);
            rootLayout.RowStyles[row].SizeType = visible ? SizeType.AutoSize : SizeType.Absolute;
            if (!visible) rootLayout.RowStyles[row].Height = 0;
            extension.Visible = visible;
        }
    }
}
