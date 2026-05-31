using System;

namespace ExcelTrainingMonitor.Models
{
    public class HistoryEntry
    {
        public string Employee { get; set; }

        public string Category { get; set; }

        public string OldStatus { get; set; }

        public string NewStatus { get; set; }

        public DateTime Timestamp { get; set; }
    }
}