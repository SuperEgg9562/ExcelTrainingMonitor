using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace ExcelTrainingMonitor.Services
{
    internal static class DocumentEditorService
    {
        public static void ResizeTextBox(TextBox textBox, int minimumHeight, int maximumHeight)
        {
            if (textBox.ClientSize.Width <= 20)
                return;

            Size measured = TextRenderer.MeasureText(
                string.IsNullOrEmpty(textBox.Text) ? " " : textBox.Text + " ",
                textBox.Font,
                new Size(textBox.ClientSize.Width - 12, int.MaxValue),
                TextFormatFlags.TextBoxControl | TextFormatFlags.WordBreak);
            int desiredHeight = measured.Height + 14;
            textBox.Height = Math.Clamp(desiredHeight, minimumHeight, maximumHeight);
            textBox.ScrollBars = desiredHeight > maximumHeight ? ScrollBars.Vertical : ScrollBars.None;
        }

        public static void ResizeDateTimePickers(DateTimePicker datePicker, DateTimePicker timePicker)
        {
            if (datePicker == null || timePicker == null)
                return;

            int dateWidth = TextRenderer.MeasureText(
                datePicker.Value.ToString("dddd, dd MMMM yyyy"), datePicker.Font).Width;
            datePicker.Width = dateWidth + SystemInformation.VerticalScrollBarWidth + 28;

            int timeWidth = TextRenderer.MeasureText(timePicker.Value.ToString("HH:mm"), timePicker.Font).Width;
            timePicker.Width = timeWidth + 34;
        }

        public static void SelectLogo(IWin32Window owner, PictureBox pictureBox, Label placeholder)
        {
            using var dialog = new OpenFileDialog
            {
                Filter = "Image Files|*.png;*.jpg;*.jpeg;*.bmp;*.gif"
            };

            if (dialog.ShowDialog(owner) != DialogResult.OK)
                return;

            using var source = Image.FromFile(dialog.FileName);
            Image logo = new Bitmap(source);
            pictureBox.Image?.Dispose();
            pictureBox.Image = logo;
            placeholder.Visible = false;
        }

        public static void ResetLogo(PictureBox pictureBox, Label placeholder)
        {
            pictureBox.Image?.Dispose();
            pictureBox.Image = null;
            placeholder.Visible = true;
        }

        public static void ResizeGrid(DataGridView grid)
        {
            if (grid.Columns.Count == 0)
                return;

            grid.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
            int columnWidth = Math.Clamp(
                grid.Columns.Cast<DataGridViewColumn>().Max(column => column.Width), 90, 320);

            foreach (DataGridViewColumn column in grid.Columns)
            {
                column.MinimumWidth = 90;
                column.AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                column.Width = columnWidth;
            }

            grid.AutoResizeRows(DataGridViewAutoSizeRowsMode.AllCells);
            int rowHeight = Math.Clamp(
                grid.Rows.Cast<DataGridViewRow>()
                    .Where(row => !row.IsNewRow)
                    .Select(row => row.Height)
                    .DefaultIfEmpty(28)
                    .Max(),
                28,
                160);

            foreach (DataGridViewRow row in grid.Rows)
                row.Height = rowHeight;
        }

        public static void DrawRowNumber(DataGridView grid, DataGridViewRowPostPaintEventArgs e, Color color)
        {
            TextRenderer.DrawText(
                e.Graphics,
                (e.RowIndex + 1).ToString(),
                grid.Font,
                new Rectangle(e.RowBounds.Left, e.RowBounds.Top, grid.RowHeadersWidth - 4, e.RowBounds.Height),
                color,
                TextFormatFlags.Right | TextFormatFlags.VerticalCenter);
        }

        public static void AddRow(DataTable table, DataGridView grid)
        {
            table.Rows.Add(table.NewRow());
            ResizeGrid(grid);
        }

        public static void AddColumn(DataTable table, DataGridView grid)
        {
            int columnNumber = table.Columns.Count;
            string columnName;
            do
            {
                columnName = GridBookEditorService.ColumnName(columnNumber++);
            }
            while (table.Columns.Contains(columnName));

            table.Columns.Add(columnName);
            ResizeGrid(grid);
        }

        public static void MoveSelectedRows(DataTable table, DataGridView grid, int direction)
        {
            grid.EndEdit();
            HashSet<int> selectedRows = GetSelectedRowIndices(table, grid);
            if (selectedRows.Count == 0)
                return;

            IEnumerable<int> orderedRows = direction < 0
                ? selectedRows.OrderBy(index => index)
                : selectedRows.OrderByDescending(index => index);

            foreach (int rowIndex in orderedRows.ToArray())
            {
                int targetIndex = rowIndex + direction;
                if (targetIndex < 0 || targetIndex >= table.Rows.Count || selectedRows.Contains(targetIndex))
                    continue;

                object[] currentValues = table.Rows[rowIndex].ItemArray;
                table.Rows[rowIndex].ItemArray = table.Rows[targetIndex].ItemArray;
                table.Rows[targetIndex].ItemArray = currentValues;
                selectedRows.Remove(rowIndex);
                selectedRows.Add(targetIndex);
            }

            SelectRows(grid, selectedRows);
            ResizeGrid(grid);
        }

        public static void ClearSelectedCells(DataGridView grid)
        {
            grid.EndEdit();
            foreach (DataGridViewCell cell in grid.SelectedCells)
            {
                if (!cell.ReadOnly && cell.RowIndex >= 0 && !grid.Rows[cell.RowIndex].IsNewRow)
                    cell.Value = "";
            }

            ResizeGrid(grid);
        }

        public static void DeleteSelectedRows(DataTable table, DataGridView grid)
        {
            grid.EndEdit();
            foreach (int rowIndex in GetSelectedRowIndices(table, grid).OrderByDescending(index => index))
            {
                if (rowIndex >= 0 && rowIndex < table.Rows.Count)
                    table.Rows.RemoveAt(rowIndex);
            }

            ResizeGrid(grid);
        }

        public static void DeleteSelectedColumns(DataTable table, DataGridView grid)
        {
            grid.EndEdit();
            HashSet<int> columnIndices = grid.SelectedCells
                .Cast<DataGridViewCell>()
                .Select(cell => cell.ColumnIndex)
                .ToHashSet();

            if (columnIndices.Count == 0 && grid.CurrentCell != null)
                columnIndices.Add(grid.CurrentCell.ColumnIndex);

            foreach (int columnIndex in columnIndices.OrderByDescending(index => index))
            {
                if (columnIndex >= 0 && columnIndex < table.Columns.Count)
                    table.Columns.RemoveAt(columnIndex);
            }

            ResizeGrid(grid);
        }

        private static HashSet<int> GetSelectedRowIndices(DataTable table, DataGridView grid)
        {
            HashSet<int> indices = grid.SelectedRows
                .Cast<DataGridViewRow>()
                .Where(row => !row.IsNewRow)
                .Select(row => row.Index)
                .ToHashSet();

            if (indices.Count == 0)
            {
                indices = grid.SelectedCells
                    .Cast<DataGridViewCell>()
                    .Where(cell => cell.RowIndex >= 0 && cell.RowIndex < table.Rows.Count)
                    .Select(cell => cell.RowIndex)
                    .ToHashSet();
            }

            return indices;
        }

        private static void SelectRows(DataGridView grid, IEnumerable<int> rowIndices)
        {
            grid.ClearSelection();
            foreach (int rowIndex in rowIndices.Where(index => index >= 0 && index < grid.Rows.Count))
            {
                foreach (DataGridViewCell cell in grid.Rows[rowIndex].Cells)
                    cell.Selected = true;
            }
        }
    }
}
