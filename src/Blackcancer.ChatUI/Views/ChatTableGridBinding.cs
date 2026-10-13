// SPDX-License-Identifier: MPL-2.0
using System;
using System.Drawing;
using System.Windows.Forms;
using Blackcancer.ChatUI.Presentation;

namespace Blackcancer.ChatUI.Views
{
    internal static class ChatTableGridBinding
    {
        internal static void Bind(DataGridView grid, ChatTableSnapshot snapshot)
        {
            grid.SuspendLayout();
            try
            {
                int dpi = Math.Max(96, grid.DeviceDpi);
                int minimumWidth = (48 * dpi + 95) / 96;
                while (grid.Columns.Count > snapshot.ColumnCount) grid.Columns.RemoveAt(grid.Columns.Count - 1);
                while (grid.Columns.Count < snapshot.ColumnCount)
                    grid.Columns.Add(new DataGridViewTextBoxColumn { SortMode = DataGridViewColumnSortMode.NotSortable, MinimumWidth = minimumWidth });
                for (int column = 0; column < snapshot.ColumnCount; column++)
                {
                    var target = grid.Columns[column];
                    target.HeaderText = snapshot.GetHeader(column);
                    var alignment = column < snapshot.ColumnAlignments.Count ? snapshot.ColumnAlignments[column] : ChatMarkdownTableAlignment.Unspecified;
                    target.DefaultCellStyle.Alignment = alignment == ChatMarkdownTableAlignment.Right ? DataGridViewContentAlignment.MiddleRight
                        : alignment == ChatMarkdownTableAlignment.Center ? DataGridViewContentAlignment.MiddleCenter : DataGridViewContentAlignment.MiddleLeft;
                    int width = Measure(target.HeaderText, grid.Font) + 24;
                    for (int row = 0; row < Math.Min(snapshot.DataRowCount, 100); row++)
                        width = Math.Max(width, Measure(snapshot.GetDataCell(row, column), grid.Font) + 20);
                    target.MinimumWidth = minimumWidth;
                    target.Width = Math.Max(minimumWidth, Math.Min(Math.Max(1, (int)Math.Ceiling(grid.Font.GetHeight(dpi))) * 36, width));
                }
                grid.RowTemplate.Height = RowHeight(grid);
                grid.ColumnHeadersHeight = Math.Max((23 * dpi + 95) / 96, (int)Math.Ceiling(grid.Font.GetHeight(dpi)) + (10 * dpi + 95) / 96);
                grid.ColumnHeadersVisible = snapshot.ColumnCount > 0;
                grid.RowCount = snapshot.ColumnCount == 0 ? 0 : snapshot.DataRowCount;
                grid.Invalidate();
            }
            finally { grid.ResumeLayout(); }
        }

        internal static int RowHeight(DataGridView grid)
        {
            int dpi = Math.Max(96, grid.DeviceDpi);
            return Math.Max((22 * dpi + 95) / 96, (int)Math.Ceiling(grid.Font.GetHeight(dpi)) + (8 * dpi + 95) / 96);
        }

        private static int Measure(string text, Font font)
        {
            text = text ?? "";
            if (text.Length > 512) text = text.Substring(0, 512);
            return TextRenderer.MeasureText(text, font, new Size(4096, font.Height), TextFormatFlags.SingleLine | TextFormatFlags.NoPadding).Width;
        }
    }
}
