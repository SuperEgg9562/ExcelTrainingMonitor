using ClosedXML.Excel;
using ExcelTrainingMonitor.Models;
using System;
using System.Collections.Generic;
using System.IO;

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
    }
}
