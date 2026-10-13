// SPDX-License-Identifier: MPL-2.0
// Ported from SolidAi native chat and VBAi native Designer-backed composer; source-vendored host-independent UI.
// Adapted from VBAi native Designer-backed composer; VBE reference selection is not included.
using System.ComponentModel;
using System.Windows.Forms;
using WpfTextBox = System.Windows.Controls.TextBox;

namespace Blackcancer.ChatUI.Views
{
    /// <summary>Saisie WinForms native éditable dans le Designer ; moteur WPF et correcteur disponibles en option.</summary>
    [ToolboxItem(true)]
    public sealed partial class ChatInputView : UserControl
    {
        /// <summary>Contrôle WPF paresseux qui reçoit le texte de saisie.</summary>
        private WpfTextBox editor;
        private bool useNativeEditor = true;
        /// <summary>État du correcteur orthographique présenté au contrôle WPF.</summary>
        private bool spellCheckEnabled = true;
        /// <summary>Marge intérieure appliquée au panneau et au moteur de texte.</summary>
        private Padding inputPadding = new Padding(12);
        /// <summary>Construit uniquement la surface Designer ; le moteur de texte est initialisé à la demande.</summary>
        public ChatInputView()
        {
            InitializeComponent();
            previewEditor.TextChanged += (sender, args) => { if (useNativeEditor || editor == null) base.OnTextChanged(System.EventArgs.Empty); };
            previewEditor.KeyDown += (sender, args) => { if (useNativeEditor) OnKeyDown(args); };
        }
        /// <summary>Uses the native WinForms editor for hosts with a native message loop.</summary>
        [DefaultValue(true), Category("Behavior")]
        public bool UseNativeEditor
        {
            get => useNativeEditor;
            set
            {
                if (value == useNativeEditor) return;
                if (value && editor != null)
                {
                    previewEditor.Text = editor.Text;
                    previewEditor.ReadOnly = editor.IsReadOnly;
                }
                else if (!value && editor != null)
                {
                    editor.Text = previewEditor.Text;
                    editor.IsReadOnly = previewEditor.ReadOnly;
                }
                useNativeEditor = value;
                if (value || DesignMode || LicenseManager.UsageMode == LicenseUsageMode.Designtime)
                {
                    host.Visible = false;
                    previewPanel.Visible = true;
                    previewPanel.BringToFront();
                }
                else if (IsHandleCreated) { var valueEditor = Editor; }
            }
        }
        [Browsable(true), DesignerSerializationVisibility(DesignerSerializationVisibility.Visible)]
        public override string Text
        {
            get => useNativeEditor ? previewEditor.Text : (editor == null ? previewEditor.Text : editor.Text);
            set
            {
                string text = value ?? string.Empty;
                if (!useNativeEditor && editor != null) editor.Text = text;
                else previewEditor.Text = text;
            }
        }
        [DefaultValue(false), Category("Behavior")]
        public bool ReadOnly
        {
            get => useNativeEditor ? previewEditor.ReadOnly : editor?.IsReadOnly ?? previewEditor.ReadOnly;
            set { previewEditor.ReadOnly = value; if (editor != null) editor.IsReadOnly = value; }
        }
        public void Clear() { Text = string.Empty; }
        [Browsable(false), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public int SelectionStart
        {
            get => useNativeEditor ? previewEditor.SelectionStart : Editor.SelectionStart;
            set { if (useNativeEditor) previewEditor.SelectionStart = value; else Editor.SelectionStart = value; }
        }
        [Browsable(false), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public int SelectionLength
        {
            get => useNativeEditor ? previewEditor.SelectionLength : Editor.SelectionLength;
            set { if (useNativeEditor) previewEditor.SelectionLength = value; else Editor.SelectionLength = value; }
        }
        public void Select(int start, int length)
        { if (useNativeEditor) previewEditor.Select(start, length); else Editor.Select(start, length); }
        protected override void OnEnter(System.EventArgs e)
        {
            base.OnEnter(e);
            FocusInput();
        }
        protected override void OnHandleCreated(System.EventArgs e)
        {
            base.OnHandleCreated(e);
            if (useNativeEditor || DesignMode || LicenseManager.UsageMode == LicenseUsageMode.Designtime) return;
            var value = Editor;
        }
        public bool FocusInput() => useNativeEditor ? previewEditor.Focus() : Editor.Focus();
        /// <summary>Compatibility entry point for a host's contextual reference marker.</summary>
        public void InsertCadMentionToken() => InsertMentionToken("#");

        public void InsertMentionToken(string token)
        {
            if (token == null) throw new System.ArgumentNullException(nameof(token));
            if (ReadOnly) return;
            int start = useNativeEditor ? previewEditor.SelectionStart : Editor.SelectionStart;
            string text = Text;
            string insertion = start > 0 && !char.IsWhiteSpace(text[start - 1]) ? " " + token : token;
            if (useNativeEditor)
            {
                previewEditor.SelectedText = insertion;
                previewEditor.Select(start + insertion.Length, 0);
            }
            else
            {
                Editor.SelectedText = insertion;
                Editor.Select(start + insertion.Length, 0);
            }
            FocusInput();
        }
                /// <summary>Moteur de texte créé uniquement lorsque le chat initialise ses comportements.</summary>
                /// <value>Contrôle WPF initialisé à la première lecture.</value>
        [Browsable(false), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public WpfTextBox Editor
        {
            get
            {
                if (useNativeEditor) throw new System.InvalidOperationException("The WPF editor is disabled while UseNativeEditor is true.");
                if (editor == null)
                {
                    editor = new WpfTextBox { AcceptsReturn = true,
                        TextWrapping = System.Windows.TextWrapping.Wrap,
                        VerticalScrollBarVisibility = System.Windows.Controls.ScrollBarVisibility.Auto,
                        BorderThickness = new System.Windows.Thickness(0),
                        Background = System.Windows.Media.Brushes.Transparent };
                    editor.Text = previewEditor.Text;
                    editor.IsReadOnly = previewEditor.ReadOnly;
                    editor.TextChanged += (sender, args) => { if (!useNativeEditor) base.OnTextChanged(System.EventArgs.Empty); };
                    editor.PreviewKeyDown += (sender, args) =>
                    {
                        var key = (Keys)System.Windows.Input.KeyInterop.VirtualKeyFromKey(args.Key);
                        var modifiers = System.Windows.Input.Keyboard.Modifiers;
                        if ((modifiers & System.Windows.Input.ModifierKeys.Shift) != 0) key |= Keys.Shift;
                        if ((modifiers & System.Windows.Input.ModifierKeys.Control) != 0) key |= Keys.Control;
                        if ((modifiers & System.Windows.Input.ModifierKeys.Alt) != 0) key |= Keys.Alt;
                        var native = new KeyEventArgs(key);
                        OnKeyDown(native);
                        args.Handled = native.Handled || native.SuppressKeyPress;
                    };
                    host.Child = editor;
                    ApplyEditorAppearance();
                    previewPanel.Visible = false;
                    host.Visible = true;
                    host.BringToFront();
                }
                previewPanel.Visible = false;
                host.Visible = true;
                host.BringToFront();
                return editor;
            }
        }
        [Browsable(false), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public WpfTextBox EditorWPF => Editor;

        [Browsable(false), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public ChatComposerTextBox NativeEditor => previewEditor;
                /// <summary>Active le correcteur WPF lorsque UseNativeEditor est faux ; la saisie native n'a pas de correcteur intégré.</summary>
                /// <value><see langword="true"/> lorsque la vérification orthographique WPF est activée.</value>
        [DefaultValue(true), Category("Behavior")]
        public bool SpellCheckEnabled
        {
            get => spellCheckEnabled;
            set { spellCheckEnabled = value; ApplyEditorAppearance(); }
        }
                /// <summary>Marges du texte dans l’éditeur, éditables dans le Designer.</summary>
                /// <value>Marge appliquée autour du texte; ses côtés doivent être non négatifs.</value>
        [Category("Layout")]
        public Padding InputPadding
        {
            get => inputPadding;
            set
            {
                if (value.Left < 0 || value.Top < 0 || value.Right < 0 || value.Bottom < 0)
                    throw new System.ArgumentOutOfRangeException(nameof(value), "Input padding cannot be negative.");
                inputPadding = value; ApplyEditorAppearance();
            }
        }
        /// <summary>Indique au Designer si la marge diffère de sa valeur par défaut.</summary>
        /// <returns><see langword="true"/> si la marge n’est plus égale à 12 pixels sur chaque côté.</returns>
        private bool ShouldSerializeInputPadding() => inputPadding != new Padding(12);
        /// <summary>Rétablit la marge intérieure par défaut de 12 pixels sur chaque côté.</summary>
        private void ResetInputPadding() => InputPadding = new Padding(12);
                /// <summary>Propage la police Designer au moteur de texte.</summary>
                /// <param name="e">Données de changement de police.</param>
        protected override void OnFontChanged(System.EventArgs e) { base.OnFontChanged(e); ApplyEditorAppearance(); }
                /// <summary>Propage la couleur Designer au moteur de texte.</summary>
                /// <param name="e">Données de changement de couleur.</param>
        protected override void OnForeColorChanged(System.EventArgs e) { base.OnForeColorChanged(e); ApplyEditorAppearance(); }
        /// <summary>Met à jour le padding, la police, la couleur et le correcteur du contrôle WPF, s’il est créé.</summary>
        private void ApplyEditorAppearance()
        {
            if (previewPanel != null) previewPanel.Padding = inputPadding;
            if (editor == null) return;
            editor.FontFamily = new System.Windows.Media.FontFamily(Font.FontFamily.Name);
            editor.FontSize = Font.SizeInPoints * 96.0 / 72.0;
            editor.FontWeight = Font.Bold ? System.Windows.FontWeights.Bold : System.Windows.FontWeights.Normal;
            editor.FontStyle = Font.Italic ? System.Windows.FontStyles.Italic : System.Windows.FontStyles.Normal;
            editor.Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(ForeColor.R, ForeColor.G, ForeColor.B));
            editor.Padding = new System.Windows.Thickness(inputPadding.Left, inputPadding.Top, inputPadding.Right, inputPadding.Bottom);
            editor.SpellCheck.IsEnabled = spellCheckEnabled;
        }
    }
}
