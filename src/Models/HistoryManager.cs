using ExcelTrainingMonitor.Models;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;

namespace ExcelTrainingMonitor
{
    internal static class HistoryManager
    {
        private static readonly string HistoryFile =
            "history.json";

        public static List<HistoryRecord> Load()
        {
            if (!File.Exists(HistoryFile))
            {
                return new List<HistoryRecord>();
            }

            string json =
                File.ReadAllText(HistoryFile);

            return JsonSerializer.Deserialize<List<HistoryRecord>>(json)
                   ?? new List<HistoryRecord>();
        }

        public static void Save(
            List<HistoryRecord> records)
        {
            string json =
                JsonSerializer.Serialize(
                    records,
                    new JsonSerializerOptions
                    {
                        WriteIndented = true
                    });

            File.WriteAllText(
                HistoryFile,
                json);
        }

        public static void Add(
            TrainingAlert alert)
        {
            List<HistoryRecord> records =
                Load();

            records.Add(
                new HistoryRecord
                {
                    EmployeeName =
                        alert.EmployeeName,

                    Category =
                        alert.Category,

                    Status =
                        alert.Status,

                    Timestamp =
                        alert.Timestamp
                });

            Save(records);
        }
    }
}