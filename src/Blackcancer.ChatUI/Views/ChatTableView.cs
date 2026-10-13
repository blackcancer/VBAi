// SPDX-License-Identifier: MPL-2.0
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Blackcancer.ChatUI.Presentation;

namespace Blackcancer.ChatUI.Views
{
    /// <summary>A bounded, horizontally scrollable native grid for Markdown tables.</summary>
    public partial class ChatTableView : ChatView
    {
        private ChatServices services = new ChatServices();
        private SharedChatStyle style = new SharedChatStyle();
        private string source;
        private string[][] values = new string[0][];
        private ChatTableSnapshot snapshot = new ChatTableSnapshot("", new string[0][], -1, null);
        private Action<string> error = value => System.Diagnostics.Debug.WriteLine(value);
        public ChatTableView()
        {
            InitializeComponent();
            grid.CellValueNeeded += (s, e) => { if (e.RowIndex >= 0 && e.RowIndex < values.Length && e.ColumnIndex < values[e.RowIndex].Length) e.Value = values[e.RowIndex][e.ColumnIndex]; };
            copy.Click += (s, e) => { try { services.CopyText?.Invoke(ToTsv()); } catch (Exception ex) { error(ex.Message); } };
            expand.Click += (s, e) =>
            {
                ChatTableInspectorForm window = null;
                try { window = CreateExpandedView(); window.Show((IWin32Window)FindForm() ?? this); }
                catch (Exception ex) { window?.Dispose(); error(ex.Message); }
            };
            BindStyle();
        }
        public DataGridView Grid => grid;
        public string SourceText => source ?? "";
        public ChatServices Services { get => services; set { if (ReferenceEquals(services, value)) return; services = value ?? new ChatServices(); BindStyle(); } }
        public SharedChatStyle Style { get => style; set { style = value ?? new SharedChatStyle(); source = null; BindStyle(); } }
        public void BindStyle()
        {
            BackColor = grid.BackgroundColor = style.Background; ForeColor = grid.ForeColor = style.Foreground; Font = style.BodyFont;
            grid.DefaultCellStyle.BackColor = style.Background; grid.DefaultCellStyle.ForeColor = style.Foreground;
            grid.ColumnHeadersDefaultCellStyle.BackColor = style.CodeBackground; grid.ColumnHeadersDefaultCellStyle.ForeColor = style.Foreground;
            grid.GridColor = style.Border; copy.Text = services.Text("Copy table"); copy.BackColor = style.Background; copy.ForeColor = style.Foreground;
            var resources = new System.ComponentModel.ComponentResourceManager(typeof(ChatTableView));
            expand.Text = services.Text(resources.GetString("expand.Text") ?? "Expand table");
            expand.BackColor = style.Background; expand.ForeColor = style.Foreground;
            MeasureGridHeight();
        }
        internal void ShowTable(ChatMarkdownBlock block, Action<string> onError)
        {
            error = onError ?? (value => System.Diagnostics.Debug.WriteLine(value));
            if (source == block.Source) return;
            int firstRow = grid.FirstDisplayedScrollingRowIndex, horizontal = grid.HorizontalScrollingOffset;
            int currentRow = grid.CurrentCell?.RowIndex ?? -1, currentColumn = grid.CurrentCell?.ColumnIndex ?? -1;
            var selected = grid.SelectedCells.Cast<DataGridViewCell>().Select(c => Tuple.Create(c.RowIndex, c.ColumnIndex)).ToArray();
            var rows = block.TableRows;
            int count = Math.Max(block.TableColumns.Count, rows.SelectMany(r => r.Cells).Select(c => c.ColumnIndex + Math.Max(1, c.ColumnSpan)).DefaultIfEmpty(0).Max());
            var cache = new string[rows.Count][];
            for (int row = 0; row < rows.Count; row++)
            {
                cache[row] = new string[count];
                foreach (var cell in rows[row].Cells)
                    if (cell.ColumnIndex >= 0 && cell.ColumnIndex < count) cache[row][cell.ColumnIndex] = string.Concat(cell.Spans.Select(span => span.Text));
            }
            int headerIndex = -1;
            for (int row = 0; row < rows.Count; row++) if (rows[row].IsHeader) { headerIndex = row; break; }
            values = cache.Where((row, index) => index != headerIndex).ToArray();
            snapshot = new ChatTableSnapshot(block.Source, cache, headerIndex, block.TableColumns.Select(column => column.Alignment));
            grid.SuspendLayout();
            try
            {
                while (grid.Columns.Count > count) grid.Columns.RemoveAt(grid.Columns.Count - 1);
                while (grid.Columns.Count < count) grid.Columns.Add(new DataGridViewTextBoxColumn { SortMode = DataGridViewColumnSortMode.NotSortable, MinimumWidth = 48 });
                for (int column = 0; column < count; column++)
                {
                    var target = grid.Columns[column];
                    target.HeaderText = headerIndex >= 0 ? cache[headerIndex][column] ?? "" : (column + 1).ToString();
                    var alignment = column < block.TableColumns.Count ? block.TableColumns[column].Alignment : ChatMarkdownTableAlignment.Unspecified;
                    target.DefaultCellStyle.Alignment = alignment == ChatMarkdownTableAlignment.Right ? DataGridViewContentAlignment.MiddleRight : alignment == ChatMarkdownTableAlignment.Center ? DataGridViewContentAlignment.MiddleCenter : DataGridViewContentAlignment.MiddleLeft;
                    // Work and width are bounded by sample size, independently of full cell length.
                    int width = MeasureCell(target.HeaderText) + 24;
                    for (int row = 0; row < Math.Min(values.Length, 100); row++) width = Math.Max(width, MeasureCell(values[row][column]) + 20);
                    int scale = Math.Max(1, Font.Height);
                    target.Width = Math.Max(48, Math.Min(scale * 36, width));
                }
                grid.ColumnHeadersVisible = count > 0;
                grid.RowCount = count == 0 ? 0 : values.Length;
                grid.ClearSelection();
                if (currentRow >= 0 && currentRow < grid.RowCount && currentColumn >= 0 && currentColumn < count) grid.CurrentCell = grid[currentColumn, currentRow];
                foreach (var cell in selected) if (cell.Item1 >= 0 && cell.Item1 < grid.RowCount && cell.Item2 >= 0 && cell.Item2 < count) grid[cell.Item2, cell.Item1].Selected = true;
                if (firstRow >= 0 && firstRow < grid.RowCount) grid.FirstDisplayedScrollingRowIndex = firstRow;
                if (count > 0) grid.HorizontalScrollingOffset = horizontal;
                source = block.Source; grid.Invalidate(); MeasureGridHeight();
            }
            finally { grid.ResumeLayout(); }
        }
        private int MeasureCell(string text)
        { text = text ?? ""; if (text.Length > 512) text = text.Substring(0, 512); return TextRenderer.MeasureText(text, Font, new Size(4096, Font.Height), TextFormatFlags.SingleLine | TextFormatFlags.NoPadding).Width; }
        private void MeasureGridHeight()
        {
            if (grid == null) return;
            int rowHeight = Math.Max(22, Font.Height + 8);
            grid.RowTemplate.Height = rowHeight;
            int height = grid.ColumnHeadersHeight + Math.Min(12, values.Length) * rowHeight + SystemInformation.HorizontalScrollBarHeight + 4;
            if (grid.Height != height) grid.Height = height;
        }
        public ChatTableSnapshot CreateSnapshot() => snapshot;

        public ChatTableInspectorForm CreateExpandedView()
        {
            var window = new ChatTableInspectorForm();
            try
            {
                window.ShowSnapshot(snapshot, services, style, error);
                window.SelectLike(grid);
                Icon icon = FindForm()?.Icon ?? services.GetWindowIcon?.Invoke();
                if (icon != null) window.UseBorrowedIcon(icon);
                return window;
            }
            catch { window.Dispose(); throw; }
        }

        public string ToTsv() => snapshot.ToTsv();
    }
}
