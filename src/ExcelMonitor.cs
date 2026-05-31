using ClosedXML.Excel;
using ExcelTrainingMonitor.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ExcelTrainingMonitor
{
    public static class ExcelMonitor
    {
        public static List<TrainingAlert> ScanFile(string path)
        {
            List<TrainingAlert> alerts =
                new List<TrainingAlert>();

            if (!File.Exists(path))
                return alerts;

            using (var workbook = new XLWorkbook(path))
            {
                if (!workbook.TryGetWorksheet("TASKS", out var sheet))
                    return alerts;

                var lastRowUsed = sheet.LastRowUsed();
                var lastColumnUsed = sheet.LastColumnUsed();

                if (lastRowUsed == null || lastColumnUsed == null)
                    return alerts;

                int lastRow =
                    lastRowUsed.RowNumber();

                int lastColumn =
                    lastColumnUsed.ColumnNumber();

                
                int headerRow = 3;

                
                int startRow = 4;

                for (int row = startRow; row <= lastRow; row++)
                {
                    string employee =
                        sheet.Cell(row, 2)
                        .GetString()
                        .Trim();

                    if (string.IsNullOrWhiteSpace(employee))
                        continue;

                    for (int col = 3; col <= lastColumn; col++)
                    {
                        var headerCell =
                            sheet.Cell(headerRow, col);

                        string category =
                            headerCell.GetFormattedString()
                            .Trim();


                        if (string.IsNullOrWhiteSpace(category))
                            continue;

                        var dataCell =
                            sheet.Cell(row, col);

                        string value =
                            dataCell.GetFormattedString()
                            .Trim();

                        string status = "";


                        if (value == "0")
                        {
                            status = "Not Trained";
                        }
                        else if (value == "1")
                        {
                            status = "In Training";
                        }
                        else if (value == "2")
                        {
                            status = "Complete";
                        }
                        else
                        {
                            continue;
                        }

                        alerts.Add(new TrainingAlert
                        {
                            EmployeeName = employee,
                            Category = category,
                            Status = status,
                            Timestamp =
                                DateTime.Now.ToString(
                                    "yyyy-MM-dd HH:mm:ss")
                        });
                    }
                      
                }
            }

            return alerts;
        }

        public static void CreateTemplate(string path)
        {
            using var workbook = new XLWorkbook();
            var sheet = workbook.Worksheets.Add("TASKS");

            sheet.Cell(3, 2).Value = "Employee";
            sheet.Cell(3, 3).Value = "General";
            sheet.Cell(4, 2).Value = "Example Employee";
            sheet.Cell(4, 3).Value = 0;
            sheet.Columns().AdjustToContents();

            workbook.SaveAs(path);
        }

        public static void SaveTrainingData(string path, IEnumerable<TrainingAlert> alerts)
        {
            XLWorkbook workbook = File.Exists(path)
                ? new XLWorkbook(path)
                : new XLWorkbook();

            using (workbook)
            {
                if (workbook.TryGetWorksheet("TASKS", out var existing))
                {
                    existing.Delete();
                }

                var sheet = workbook.Worksheets.Add("TASKS");
                WriteTaskMatrix(sheet, alerts);
                workbook.SaveAs(path);
            }
        }

        public static void ExportFlatWorkbook(string path, IEnumerable<TrainingAlert> alerts)
        {
            using var workbook = new XLWorkbook();
            var sheet = workbook.Worksheets.Add("Training");

            sheet.Cell(1, 1).Value = "Employee";
            sheet.Cell(1, 2).Value = "Category";
            sheet.Cell(1, 3).Value = "Status";
            sheet.Cell(1, 4).Value = "Timestamp";

            int row = 2;
            foreach (TrainingAlert alert in alerts)
            {
                sheet.Cell(row, 1).Value = alert.EmployeeName;
                sheet.Cell(row, 2).Value = alert.Category;
                sheet.Cell(row, 3).Value = alert.Status;
                sheet.Cell(row, 4).Value = alert.Timestamp;
                row++;
            }

            sheet.Columns().AdjustToContents();
            workbook.SaveAs(path);
        }

        private static void WriteTaskMatrix(IXLWorksheet sheet, IEnumerable<TrainingAlert> alerts)
        {
            List<TrainingAlert> cleanAlerts = alerts
                .Where(x =>
                    !string.IsNullOrWhiteSpace(x.EmployeeName) &&
                    !string.IsNullOrWhiteSpace(x.Category) &&
                    !string.IsNullOrWhiteSpace(x.Status))
                .ToList();

            List<string> employees = cleanAlerts
                .Select(x => x.EmployeeName.Trim())
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(x => x)
                .ToList();

            List<string> categories = cleanAlerts
                .Select(x => x.Category.Trim())
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(x => x)
                .ToList();

            sheet.Cell(3, 2).Value = "Employee";
            for (int i = 0; i < categories.Count; i++)
            {
                sheet.Cell(3, i + 3).Value = categories[i];
            }

            for (int rowIndex = 0; rowIndex < employees.Count; rowIndex++)
            {
                string employee = employees[rowIndex];
                int row = rowIndex + 4;
                sheet.Cell(row, 2).Value = employee;

                for (int colIndex = 0; colIndex < categories.Count; colIndex++)
                {
                    string category = categories[colIndex];
                    TrainingAlert match = cleanAlerts.FirstOrDefault(x =>
                        string.Equals(x.EmployeeName.Trim(), employee, StringComparison.OrdinalIgnoreCase) &&
                        string.Equals(x.Category.Trim(), category, StringComparison.OrdinalIgnoreCase));

                    if (match != null)
                    {
                        sheet.Cell(row, colIndex + 3).Value = StatusToValue(match.Status);
                    }
                }
            }

            sheet.Columns().AdjustToContents();
        }

        private static int StatusToValue(string status)
        {
            return status switch
            {
                "Not Trained" => 0,
                "In Training" => 1,
                "Complete" => 2,
                _ => 0
            };
        }
    }
}
