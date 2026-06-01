using ClosedXML.Excel;
using System;
using System.Data;
using System.IO;
using System.Linq;

namespace ExcelTrainingMonitor.Services
{
    internal static class GridBookEditorService
    {
        public static void CreateBlankGridBook(string path)
        {
            using var workbook = new XLWorkbook();
            var sheet = workbook.Worksheets.Add("Sheet1");
            sheet.Cell(1, 1).Value = "";
            workbook.SaveAs(path);
        }

        public static string[] GetSheetNames(string path)
        {
            if (!File.Exists(path))
                return Array.Empty<string>();

            using var workbook = new XLWorkbook(path);
            return workbook.Worksheets.Select(x => x.Name).ToArray();
        }

        public static DataTable LoadSheet(string path, string sheetName, int minimumRows = 50, int minimumColumns = 26)
        {
            var table = CreateTable(minimumColumns);

            if (!File.Exists(path))
                return table;

            using var workbook = new XLWorkbook(path);
            if (!workbook.TryGetWorksheet(sheetName, out var sheet))
                return table;

            int rows = Math.Max(minimumRows, sheet.LastRowUsed()?.RowNumber() ?? minimumRows);
            int columns = Math.Max(minimumColumns, sheet.LastColumnUsed()?.ColumnNumber() ?? minimumColumns);
            table = CreateTable(columns);

            for (int row = 1; row <= rows; row++)
            {
                DataRow dataRow = table.NewRow();
                for (int column = 1; column <= columns; column++)
                {
                    dataRow[column - 1] = sheet.Cell(row, column).GetFormattedString();
                }

                table.Rows.Add(dataRow);
            }

            return table;
        }

        public static void SaveSheet(string path, string sheetName, DataTable table)
        {
            XLWorkbook workbook = File.Exists(path)
                ? new XLWorkbook(path)
                : new XLWorkbook();

            using (workbook)
            {
                if (workbook.TryGetWorksheet(sheetName, out var existing))
                {
                    existing.Delete();
                }

                var sheet = workbook.Worksheets.Add(string.IsNullOrWhiteSpace(sheetName) ? "Sheet1" : sheetName);

                for (int row = 0; row < table.Rows.Count; row++)
                {
                    for (int column = 0; column < table.Columns.Count; column++)
                    {
                        string value = table.Rows[row][column]?.ToString() ?? "";
                        sheet.Cell(row + 1, column + 1).Value = value;
                    }
                }

                sheet.Columns().AdjustToContents();
                workbook.SaveAs(path);
            }
        }

        public static void AddSheet(string path, string sheetName)
        {
            XLWorkbook workbook = File.Exists(path)
                ? new XLWorkbook(path)
                : new XLWorkbook();

            using (workbook)
            {
                string finalName = string.IsNullOrWhiteSpace(sheetName) ? "Sheet1" : sheetName.Trim();
                if (workbook.Worksheets.Any(x => string.Equals(x.Name, finalName, StringComparison.OrdinalIgnoreCase)))
                {
                    int i = 2;
                    while (workbook.Worksheets.Any(x => string.Equals(x.Name, $"{finalName}{i}", StringComparison.OrdinalIgnoreCase)))
                    {
                        i++;
                    }

                    finalName = $"{finalName}{i}";
                }

                workbook.Worksheets.Add(finalName);
                workbook.SaveAs(path);
            }
        }

        public static void ExportGridBook(string sourcePath, string destinationPath)
        {
            if (!File.Exists(sourcePath))
                throw new FileNotFoundException("GridBook not found.", sourcePath);

            File.Copy(sourcePath, destinationPath, true);
        }

        public static string ColumnName(int index)
        {
            string name = "";
            int value = index + 1;

            while (value > 0)
            {
                value--;
                name = (char)('A' + value % 26) + name;
                value /= 26;
            }

            return name;
        }

        private static DataTable CreateTable(int columns)
        {
            var table = new DataTable();
            for (int i = 0; i < columns; i++)
            {
                table.Columns.Add(ColumnName(i));
            }

            return table;
        }
    }
}
