using System;

namespace ExcelTrainingMonitor.Models
{
    public class AppSettings
    {
        public string ExcelPath { get; set; } = "";
        public int ScanIntervalHours { get; set; } = 0;
        public int ScanIntervalMinutes { get; set; } = 5;
        public bool ReminderEnabled { get; set; } = false;
        public DateTime ReminderDateTime { get; set; } = DateTime.Today.AddDays(1).AddHours(9);
        public string ThemeName { get; set; } = "Dark";
    }
}
