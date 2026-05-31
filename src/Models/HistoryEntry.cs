using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTrainingMonitor.Models
{
    public class HistoryEntry
    {
        public string Employee { get; set; }

        public string Category { get; set; }

        public string OldStatus { get; set; }

        public string NewStatus { get; set; }

        public string Timestamp { get; set; }
    }
}