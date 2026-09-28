using System.ComponentModel;
using System.Windows.Forms;
using WpfTextBox = System.Windows.Controls.TextBox;

namespace CodexVBE
{
    /// <summary>Vue WinForms éditable du moteur de saisie avec correcteur orthographique.</summary>
    [ToolboxItem(true)]
    public sealed partial class ChatInputView : UserControl
    {
        private WpfTextBox editor;
        private bool spellCheckEnabled = true;
        private Padding inputPadding = new Padding(12);
        /// <summary>Construit uniquement la surface Designer ; le moteur de texte est initialisé à la demande.</summary>
        public ChatInputView()
        {
            InitializeComponent();
        }
        /// <summary>Moteur de texte créé uniquement lorsque le chat initialise ses comportements.</summary>
        [Browsable(false), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        internal WpfTextBox Editor
        {
            get
            {
                if (editor == null)
                {
                    editor = new WpfTextBox { AcceptsReturn = true,
                        TextWrapping = System.Windows.TextWrapping.Wrap,
                        VerticalScrollBarVisibility = System.Windows.Controls.ScrollBarVisibility.Auto,
                        BorderThickness = new System.Windows.Thickness(0),
                        Background = System.Windows.Media.Brushes.Transparent };
                    host.Child = editor;
                    ApplyEditorAppearance();
                    previewPanel.Visible = false;
                    host.Visible = true;
                    host.BringToFront();
                }
                return editor;
            }
        }
        /// <summary>Active le correcteur intégré sans modifier la structure du formulaire.</summary>
        [DefaultValue(true), Category("Behavior")]
        public bool SpellCheckEnabled
        {
            get => spellCheckEnabled;
            set { spellCheckEnabled = value; ApplyEditorAppearance(); }
        }
        /// <summary>Marges du texte dans l’éditeur, éditables dans le Designer.</summary>
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
        private bool ShouldSerializeInputPadding() => inputPadding != new Padding(12);
        private void ResetInputPadding() => InputPadding = new Padding(12);
        /// <summary>Propage la police Designer au moteur de texte.</summary>
        protected override void OnFontChanged(System.EventArgs e) { base.OnFontChanged(e); ApplyEditorAppearance(); }
        /// <summary>Propage la couleur Designer au moteur de texte.</summary>
        protected override void OnForeColorChanged(System.EventArgs e) { base.OnForeColorChanged(e); ApplyEditorAppearance(); }
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
