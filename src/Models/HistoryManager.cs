using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;

namespace ExcelTrainingMonitor.Models
{
    internal static class HistoryManager
    {
        private static readonly string HistoryDirectory =
            Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "ExcelTrainingMonitor");

        private static readonly string HistoryFile =
            Path.Combine(
                HistoryDirectory,
                "History.json");

        private static List<HistoryEntry> history =
            LoadHistory();

        public static void Add(
            string employee,
            string category,
            string oldStatus,
            string newStatus)
        {
            history.Add(new HistoryEntry
            {
                Employee = employee,
                Category = category,
                OldStatus = oldStatus,
                NewStatus = newStatus,
                Timestamp = DateTime.Now
            });

            SaveHistory();
        }

        public static List<HistoryEntry> GetHistory()
        {
            return new List<HistoryEntry>(history);
        }

        public static void Clear()
        {
            history.Clear();
            SaveHistory();
        }

        private static void SaveHistory()
        {
            try
            {
                string json =
                    JsonSerializer.Serialize(
                        history,
                        new JsonSerializerOptions
                        {
                            WriteIndented = true
                        });

                Directory.CreateDirectory(HistoryDirectory);

                File.WriteAllText(
                    HistoryFile,
                    json);
            }
            catch
            {
            }
        }

        private static List<HistoryEntry> LoadHistory()
        {
            try
            {
                if (!File.Exists(HistoryFile))
                {
                    return new List<HistoryEntry>();
                }

                string json =
                    File.ReadAllText(HistoryFile);

                return JsonSerializer.Deserialize<List<HistoryEntry>>(json)
                       ?? new List<HistoryEntry>();
            }
            catch
            {
                return new List<HistoryEntry>();
            }
        }
    }
}
