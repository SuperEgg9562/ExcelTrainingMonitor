using ClosedXML.Excel;
using ExcelTrainingMonitor.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Windows.Forms;

namespace ExcelTrainingMonitor.Services
{
    internal static class ProcessRecordGridService
    {
        private const string MetadataSheetName = "Process Record Metadata";

        public static ProcessRecordMetadata LoadMetadata(string path)
        {
            if (!File.Exists(path))
                return new ProcessRecordMetadata();

            try
            {
                using var workbook = new XLWorkbook(path);
                if (!workbook.TryGetWorksheet(MetadataSheetName, out var sheet))
                    return new ProcessRecordMetadata();

                string json = sheet.Cell(1, 1).GetString();
                ProcessRecordMetadata metadata =
                    JsonSerializer.Deserialize<ProcessRecordMetadata>(json) ?? new ProcessRecordMetadata();
                metadata.DropdownLists = new Dictionary<string, List<string>>(
                    metadata.DropdownLists ?? new Dictionary<string, List<string>>(),
                    StringComparer.OrdinalIgnoreCase);
                metadata.CellDropdownAssignments ??= new Dictionary<string, string>();
                return metadata;
            }
            catch
            {
                return new ProcessRecordMetadata();
            }
        }

        public static void SaveMetadata(string path, ProcessRecordMetadata metadata)
        {
            using var workbook = File.Exists(path) ? new XLWorkbook(path) : new XLWorkbook();
            if (workbook.TryGetWorksheet(MetadataSheetName, out var existing))
                existing.Delete();

            var sheet = workbook.Worksheets.Add(MetadataSheetName);
            sheet.Cell(1, 1).Value = JsonSerializer.Serialize(metadata);
            sheet.Hide();
            workbook.SaveAs(path);
        }

        public static void ApplyAssignments(DataGridView grid, ProcessRecordMetadata metadata)
        {
            foreach (var assignment in metadata.CellDropdownAssignments.ToArray())
            {
                if (!TryParseCellKey(assignment.Key, out int rowIndex, out int columnIndex) ||
                    rowIndex < 0 || rowIndex >= grid.Rows.Count ||
                    columnIndex < 0 || columnIndex >= grid.Columns.Count ||
                    !metadata.DropdownLists.TryGetValue(assignment.Value, out List<string> items))
                {
                    continue;
                }

                SetDropdownCell(grid.Rows[rowIndex].Cells[columnIndex], assignment.Value, items);
            }
        }

        public static void AssignSelectedCells(
            DataGridView grid,
            string listName,
            IReadOnlyCollection<string> items)
        {
            foreach (DataGridViewCell cell in grid.SelectedCells)
            {
                if (cell.RowIndex >= 0 && !grid.Rows[cell.RowIndex].IsNewRow)
                    SetDropdownCell(cell, listName, items);
            }
        }

        public static void RemoveDropdownsFromSelectedCells(DataGridView grid)
        {
            foreach (DataGridViewCell cell in grid.SelectedCells.Cast<DataGridViewCell>().ToArray())
            {
                if (cell is not DataGridViewComboBoxCell)
                    continue;

                object value = cell.Value;
                var replacement = new DataGridViewTextBoxCell { Value = value, Tag = null };
                grid.Rows[cell.RowIndex].Cells[cell.ColumnIndex] = replacement;
            }
        }

        public static void UpdateListCells(
            DataGridView grid,
            string listName,
            IReadOnlyCollection<string> items)
        {
            foreach (DataGridViewRow row in grid.Rows)
            {
                foreach (DataGridViewCell cell in row.Cells.Cast<DataGridViewCell>().ToArray())
                {
                    if (cell is DataGridViewComboBoxCell &&
                        string.Equals(cell.Tag as string, listName, StringComparison.OrdinalIgnoreCase))
                    {
                        SetDropdownCell(cell, listName, items);
                    }
                }
            }
        }

        public static void RemoveListCells(DataGridView grid, string listName)
        {
            foreach (DataGridViewRow row in grid.Rows)
            {
                foreach (DataGridViewCell cell in row.Cells.Cast<DataGridViewCell>().ToArray())
                {
                    if (cell is DataGridViewComboBoxCell &&
                        string.Equals(cell.Tag as string, listName, StringComparison.OrdinalIgnoreCase))
                    {
                        object value = cell.Value;
                        row.Cells[cell.ColumnIndex] = new DataGridViewTextBoxCell
                        {
                            Value = value,
                            Tag = null
                        };
                    }
                }
            }
        }

        public static Dictionary<string, string> CaptureAssignments(DataGridView grid)
        {
            var assignments = new Dictionary<string, string>();
            foreach (DataGridViewRow row in grid.Rows)
            {
                if (row.IsNewRow)
                    continue;

                foreach (DataGridViewCell cell in row.Cells)
                {
                    if (cell is DataGridViewComboBoxCell && cell.Tag is string listName)
                        assignments[CellKey(cell.RowIndex, cell.ColumnIndex)] = listName;
                }
            }

            return assignments;
        }

        private static void SetDropdownCell(
            DataGridViewCell sourceCell,
            string listName,
            IEnumerable<string> items)
        {
            object currentValue = sourceCell.Value;
            if (currentValue == null || currentValue == DBNull.Value)
                currentValue = "";
            var values = items
                .Where(item => !string.IsNullOrWhiteSpace(item))
                .Select(item => item.Trim())
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
            values.Insert(0, "");

            string currentText = currentValue?.ToString() ?? "";
            if (!string.IsNullOrWhiteSpace(currentText) &&
                !values.Contains(currentText, StringComparer.OrdinalIgnoreCase))
            {
                values.Add(currentText);
            }

            var dropdown = new DataGridViewComboBoxCell
            {
                DisplayStyle = DataGridViewComboBoxDisplayStyle.DropDownButton,
                FlatStyle = FlatStyle.Flat,
                Tag = listName
            };
            dropdown.Items.AddRange(values.Cast<object>().ToArray());
            dropdown.Value = currentValue;
            sourceCell.DataGridView.Rows[sourceCell.RowIndex].Cells[sourceCell.ColumnIndex] = dropdown;
        }

        private static string CellKey(int rowIndex, int columnIndex) => $"{rowIndex}:{columnIndex}";

        private static bool TryParseCellKey(string key, out int rowIndex, out int columnIndex)
        {
            rowIndex = -1;
            columnIndex = -1;
            string[] parts = (key ?? "").Split(':');
            return parts.Length == 2 &&
                int.TryParse(parts[0], out rowIndex) &&
                int.TryParse(parts[1], out columnIndex);
        }
    }
}
