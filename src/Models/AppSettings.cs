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
        public bool NtfyEnabled { get; set; } = false;
        public string NtfyTopic { get; set; } = "";
        public string NtfyEmail { get; set; } = "";
        public string MinimizeBehavior { get; set; } = "Tray";
        public bool StartReminderAgentOnClose { get; set; } = true;
        public DateTime LastAgentReminderSentFor { get; set; } = DateTime.MinValue;
    }
}
