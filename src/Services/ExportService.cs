using ExcelTrainingMonitor.Models;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;

namespace ExcelTrainingMonitor.Services
{
    internal static class ExportService
    {
        public static void ExportAlertsCsv(string path, IEnumerable<TrainingAlert> alerts)
        {
            var builder = new StringBuilder();
            builder.AppendLine("Employee,Category,Status,Timestamp");

            foreach (TrainingAlert alert in alerts)
            {
                builder.AppendCsv(alert.EmployeeName);
                builder.Append(',');
                builder.AppendCsv(alert.Category);
                builder.Append(',');
                builder.AppendCsv(alert.Status);
                builder.Append(',');
                builder.AppendCsv(alert.Timestamp);
                builder.AppendLine();
            }

            File.WriteAllText(path, builder.ToString(), Encoding.UTF8);
        }

        public static void ExportHistoryCsv(string path, IEnumerable<HistoryEntry> history)
        {
            var builder = new StringBuilder();
            builder.AppendLine("Employee,Category,Old Status,New Status,Timestamp");

            foreach (HistoryEntry entry in history)
            {
                builder.AppendCsv(entry.Employee);
                builder.Append(',');
                builder.AppendCsv(entry.Category);
                builder.Append(',');
                builder.AppendCsv(entry.OldStatus);
                builder.Append(',');
                builder.AppendCsv(entry.NewStatus);
                builder.Append(',');
                builder.AppendCsv(entry.Timestamp.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture));
                builder.AppendLine();
            }

            File.WriteAllText(path, builder.ToString(), Encoding.UTF8);
        }

        private static void AppendCsv(this StringBuilder builder, string value)
        {
            value ??= "";
            bool quote = value.Contains(',') || value.Contains('"') || value.Contains('\r') || value.Contains('\n');
            string escaped = value.Replace("\"", "\"\"");

            if (quote)
            {
                builder.Append('"');
                builder.Append(escaped);
                builder.Append('"');
                return;
            }

            builder.Append(escaped);
        }
    }
}
