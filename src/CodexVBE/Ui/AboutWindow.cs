using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Reflection;
using System.Windows.Forms;

namespace CodexVBE
{
    /// <summary>Présente l’identité VBAi, sa version et les ressources de support.</summary>
    internal sealed partial class AboutWindow : Form
    {
        internal Action<string> CopyText = Clipboard.SetText;
        internal Action<string> OpenLink = SafeLinks.Open;
        internal static Func<Assembly> MetadataAssembly = ReadMetadataAssembly;
        internal static Func<bool> ProcessIs64Bit = ReadProcessIs64Bit;
        internal static Func<Version> RuntimeVersion = ReadRuntimeVersion;
        private static Version ReadRuntimeVersion() => Environment.Version;
        private static Assembly ReadMetadataAssembly() => typeof(AboutWindow).Assembly;
        private static bool ReadProcessIs64Bit() => Environment.Is64BitProcess;
        internal static Assembly ResolveImageReader(object sender, ResolveEventArgs request) =>
            request.Name == "System.Resources.Extensions, Version=4.0.0.0, Culture=neutral, PublicKeyToken=cc7b13ffcd2ddd51"
                ? typeof(System.Resources.Extensions.DeserializingResourceReader).Assembly : null;
        private string hostProcess;
        private bool runtimeInitialized;

        /// <summary>Construit la fenêtre Designer puis renseigne les métadonnées locales à l’exécution.</summary>
        public AboutWindow()
        {
            // SDK-generated image resources name the legacy reader assembly. COM hosts
            // and the Designer do not use this library's binding-redirect configuration.
            // Resolve only that reader while loading this window's resources.
            ResolveEventHandler imageReader = ResolveImageReader;
            AppDomain.CurrentDomain.AssemblyResolve += imageReader;
            try { InitializeComponent(); }
            finally { AppDomain.CurrentDomain.AssemblyResolve -= imageReader; }
            if (LicenseManager.UsageMode == LicenseUsageMode.Designtime) return;
            Icon = VbeWindowIcons.Icon("assistant");
            UiText.Apply(this, components);
            var assembly = MetadataAssembly();
            var information = assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>();
            versionValue.Text = information?.InformationalVersion ?? assembly.GetName().Version.ToString();
            using (var process = Process.GetCurrentProcess()) hostProcess = process.ProcessName;
            hostValue.Text = HostDescription(hostProcess);
            platformValue.Text = "Windows · " + (ProcessIs64Bit() ? "x64" : "x86") + " · .NET Framework 4.8";
            languageValue.Text = UiText.Culture.NativeName;
            runtimeInitialized = true;
            ApplyAppearance();
            UiTheme.Changed += ApplyAppearance;
        }

        /// <summary>Informations techniques copiables, sans chemins, identifiants ou contenu de projet.</summary>
        internal string TechnicalDetails => "VBAi " + versionValue.Text + Environment.NewLine +
            "Host: " + hostProcess + Environment.NewLine +
            "Platform: " + platformValue.Text + Environment.NewLine +
            "CLR: " + RuntimeVersion() + Environment.NewLine +
            "Interface: " + UiText.Culture.Name + Environment.NewLine +
            "Theme: " + UiTheme.Choice;

        /// <summary>Décrit les hôtes connus et conserve le nom de processus pour les autres hôtes.</summary>
        internal static string HostDescription(string processName)
        {
            if (string.Equals(processName, "EXCEL", StringComparison.OrdinalIgnoreCase)) return "Microsoft Excel · Visual Basic Editor";
            if (string.Equals(processName, "SLDWORKS", StringComparison.OrdinalIgnoreCase)) return "SOLIDWORKS · Visual Basic Editor";
            return processName;
        }

        private void ApplyAppearance()
        {
            if (IsDisposed) return;
            if (InvokeRequired) { BeginInvoke(new Action(ApplyAppearance)); return; }
            var linkColor = UiTheme.HighContrast() ? SystemColors.HotTrack :
                UiTheme.Dark ? Color.FromArgb(147, 197, 253) : Color.FromArgb(29, 78, 216);
            projectLink.LinkColor = projectLink.ActiveLinkColor = projectLink.VisitedLinkColor = linkColor;
            documentationLink.LinkColor = documentationLink.ActiveLinkColor = documentationLink.VisitedLinkColor = linkColor;
            issuesLink.LinkColor = issuesLink.ActiveLinkColor = issuesLink.VisitedLinkColor = linkColor;
            detailsLayout.BackColor = UiTheme.Surface;
            versionLabel.BackColor = versionValue.BackColor = hostLabel.BackColor = hostValue.BackColor =
                platformLabel.BackColor = platformValue.BackColor = languageLabel.BackColor = languageValue.BackColor = UiTheme.Surface;
            closeButton.BackColor = UiTheme.HighContrast() ? SystemColors.Highlight : UiTheme.Dark ? Color.FromArgb(37, 99, 235) : Color.FromArgb(29, 78, 216);
            closeButton.ForeColor = UiTheme.HighContrast() ? SystemColors.HighlightText : Color.White;
        }

        private void Updates_Click(object sender, EventArgs e)
        {
            using (var window = new UpdateWindow()) AddIn.ShowModal(window, this);
        }

        private void CopyDetails_Click(object sender, EventArgs e)
        {
            try { CopyText(TechnicalDetails); status.Text = UiText.Get("Technical details copied."); }
            catch (Exception error) { LoadLog.Write("About details copy failed: " + error.Message); status.Text = UiText.Get("Unable to copy technical details."); }
        }

        private void ResourceLink_Click(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try { OpenLink((string)((LinkLabel)sender).Tag); status.Text = ""; }
            catch (Exception error) { LoadLog.Write("About resource link failed: " + error.Message); status.Text = UiText.Get("Unable to open the link."); }
        }

        private void DisposeRuntime()
        {
            if (runtimeInitialized) UiTheme.Changed -= ApplyAppearance;
        }

        /// <summary>Ouvre À propos depuis le VBE, même si le chat est fermé.</summary>
        internal static void ShowForVbe(object vbe)
        {
            IWin32Window owner = null;
            try { owner = new NativeOwner(new IntPtr(Convert.ToInt64(((dynamic)vbe).MainWindow.HWnd))); }
            catch (Exception error) { LoadLog.Write("About VBE owner unavailable: " + error.Message); }
            using (var window = new AboutWindow()) AddIn.ShowModal(window, owner);
        }
        private sealed class NativeOwner : IWin32Window
        {
            internal NativeOwner(IntPtr handle) { Handle = handle; }
            public IntPtr Handle { get; }
        }
    }
}
