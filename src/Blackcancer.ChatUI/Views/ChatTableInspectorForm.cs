// SPDX-License-Identifier: MPL-2.0
using System;
using System.ComponentModel;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;
using Blackcancer.ChatUI.Presentation;

namespace Blackcancer.ChatUI.Views
{
    /// <summary>Independent reading window for one immutable table snapshot.</summary>
    public partial class ChatTableInspectorForm : Form
    {
        private ChatTableSnapshot snapshot = new ChatTableSnapshot("", new string[0][], -1, null);
        private ChatServices services = new ChatServices();
        private Action<string> error = value => System.Diagnostics.Debug.WriteLine(value);
        private Font readingFont;
        private Font layoutFont;
        private Icon borrowedIconClone;
        private bool binding;
        private string fontFamily;
        private FontStyle fontStyle;

        public ChatTableInspectorForm()
        {
            InitializeComponent();
            grid.CellValueNeeded += (sender, args) => args.Value = snapshot.GetDataCell(args.RowIndex, args.ColumnIndex);
            grid.RowHeightInfoNeeded += (sender, args) => { args.Height = ChatTableGridBinding.RowHeight(grid); args.MinimumHeight = args.Height; };
            grid.SelectionChanged += (sender, args) => ShowSelectedCell();
            textSize.ValueChanged += (sender, args) => { if (!binding) ChangeReadingSize((float)textSize.Value); };
            copy.Click += (sender, args) => { try { services.CopyText?.Invoke(snapshot.ToTsv()); } catch (Exception exception) { error(exception.Message); } };
            close.Click += (sender, args) => Close();
        }

        [Browsable(false), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public ChatTableSnapshot Snapshot => snapshot;
        [Browsable(false), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public DataGridView Grid => grid;
        [Browsable(false), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public TextBox CellContent => selectedCell;

        public void ShowSnapshot(ChatTableSnapshot value, ChatServices callbacks = null, SharedChatStyle style = null, Action<string> onError = null)
        {
            snapshot = value ?? throw new ArgumentNullException(nameof(value));
            services = callbacks ?? new ChatServices();
            error = onError ?? (message => System.Diagnostics.Debug.WriteLine(message));
            style = style ?? new SharedChatStyle();
            var resources = new ComponentResourceManager(typeof(ChatTableInspectorForm));
            Text = services.Text(resources.GetString("$this.Text") ?? "Table snapshot");
            copy.Text = services.Text(resources.GetString("copy.Text") ?? "Copy table");
            close.Text = services.Text(resources.GetString("close.Text") ?? "Close");
            textSizeLabel.Text = services.Text(resources.GetString("textSizeLabel.Text") ?? "Text size");
            cellLabel.Text = services.Text(resources.GetString("cellLabel.Text") ?? "Selected cell text");
            note.Text = services.Text(resources.GetString("note.Text") ?? "Captured text, including units and unknown values, is preserved.");
            summary.Text = string.Format(CultureInfo.CurrentCulture,
                services.Text(resources.GetString("summaryFormat") ?? "{0} rows · {1} columns · captured snapshot"), snapshot.DataRowCount, snapshot.ColumnCount);
            BackColor = layout.BackColor = toolbar.BackColor = detailLayout.BackColor = style.Background;
            ForeColor = style.Foreground;
            grid.BackgroundColor = grid.DefaultCellStyle.BackColor = selectedCell.BackColor = style.Background;
            grid.ForeColor = grid.DefaultCellStyle.ForeColor = selectedCell.ForeColor = style.Foreground;
            textSize.BackColor = style.Background; textSize.ForeColor = style.Foreground;
            grid.ColumnHeadersDefaultCellStyle.BackColor = style.CodeBackground;
            grid.ColumnHeadersDefaultCellStyle.ForeColor = style.Foreground;
            grid.GridColor = style.Border;
            foreach (Button button in new[] { copy, close }) { button.BackColor = style.Background; button.ForeColor = style.Foreground; }
            RightToLeft = style.RightToLeft ? RightToLeft.Yes : RightToLeft.No;
            RightToLeftLayout = style.RightToLeft;
            Font original = style.BodyFont ?? SystemFonts.MessageBoxFont;
            Font previousLayout = layoutFont;
            var nextLayout = (Font)original.Clone();
            if (layoutFont != null && layoutFont.Equals(nextLayout)) nextLayout.Dispose();
            else { Font = layoutFont = nextLayout; previousLayout?.Dispose(); }
            fontFamily = original.FontFamily.Name; fontStyle = original.Style;
            binding = true;
            try { textSize.Value = Math.Max(textSize.Minimum, Math.Min(textSize.Maximum, (decimal)original.SizeInPoints)); }
            finally { binding = false; }
            ChangeReadingSize((float)textSize.Value);
            copy.Enabled = snapshot.Rows.Count > 0;
            ShowSelectedCell();
        }

        internal void UseBorrowedIcon(Icon icon)
        {
            Icon next = (Icon)icon.Clone();
            Icon previous = borrowedIconClone;
            Icon = borrowedIconClone = next;
            previous?.Dispose();
        }

        internal void SelectLike(DataGridView source)
        {
            var selected = source.SelectedCells.Cast<DataGridViewCell>().Select(cell => Tuple.Create(cell.RowIndex, cell.ColumnIndex)).ToArray();
            var current = source.CurrentCell;
            grid.ClearSelection();
            if (current != null && current.RowIndex < grid.RowCount && current.ColumnIndex < grid.ColumnCount)
                grid.CurrentCell = grid[current.ColumnIndex, current.RowIndex];
            foreach (var cell in selected)
                if (cell.Item1 >= 0 && cell.Item1 < grid.RowCount && cell.Item2 >= 0 && cell.Item2 < grid.ColumnCount)
                    grid[cell.Item2, cell.Item1].Selected = true;
            int firstRow = source.FirstDisplayedScrollingRowIndex;
            if (firstRow >= 0 && firstRow < grid.RowCount) grid.FirstDisplayedScrollingRowIndex = firstRow;
            if (grid.ColumnCount > 0) grid.HorizontalScrollingOffset = source.HorizontalScrollingOffset;
            ShowSelectedCell();
        }

        private void ChangeReadingSize(float size)
        {
            int row = grid.CurrentCell?.RowIndex ?? -1, column = grid.CurrentCell?.ColumnIndex ?? -1;
            int firstRow = grid.FirstDisplayedScrollingRowIndex, horizontal = grid.HorizontalScrollingOffset;
            var selected = grid.SelectedCells.Cast<DataGridViewCell>().Select(cell => Tuple.Create(cell.RowIndex, cell.ColumnIndex)).ToArray();
            int selectionStart = selectedCell.SelectionStart, selectionLength = selectedCell.SelectionLength;
            var next = new Font(fontFamily ?? Font.FontFamily.Name, size, fontStyle, GraphicsUnit.Point);
            Font previous = readingFont;
            if (previous != null && previous.Equals(next)) { next.Dispose(); next = previous; previous = null; }
            grid.Font = selectedCell.Font = readingFont = next;
            ChatTableGridBinding.Bind(grid, snapshot);
            grid.ClearSelection();
            if (row >= 0 && row < grid.RowCount && column >= 0 && column < grid.ColumnCount) grid.CurrentCell = grid[column, row];
            foreach (var cell in selected)
                if (cell.Item1 < grid.RowCount && cell.Item2 < grid.ColumnCount) grid[cell.Item2, cell.Item1].Selected = true;
            if (firstRow >= 0 && firstRow < grid.RowCount) grid.FirstDisplayedScrollingRowIndex = firstRow;
            if (grid.ColumnCount > 0) grid.HorizontalScrollingOffset = horizontal;
            ShowSelectedCell();
            selectedCell.Select(Math.Min(selectionStart, selectedCell.TextLength), Math.Min(selectionLength, Math.Max(0, selectedCell.TextLength - selectionStart)));
            previous?.Dispose();
        }

        private void ShowSelectedCell()
        {
            var current = grid.CurrentCell;
            string text = current == null ? "" : snapshot.GetDataCell(current.RowIndex, current.ColumnIndex);
            if (selectedCell.Text != text) selectedCell.Text = text;
        }

        private void DisposeOwnedAppearance()
        {
            readingFont?.Dispose(); readingFont = null;
            layoutFont?.Dispose(); layoutFont = null;
            borrowedIconClone?.Dispose(); borrowedIconClone = null;
        }
    }
}
