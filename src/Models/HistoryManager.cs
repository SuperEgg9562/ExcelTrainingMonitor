using ExcelTrainingMonitor.Models;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;

namespace ExcelTrainingMonitor.Models
{
    public static class HistoryManager
    {
        private static List<HistoryEntry> history =
            new List<HistoryEntry>();

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
                Timestamp =
                    System.DateTime.Now.ToString(
                        "yyyy-MM-dd HH:mm:ss")
            });
        }

        public static List<HistoryEntry> GetHistory()
        {
            return history;
        }
    }
}