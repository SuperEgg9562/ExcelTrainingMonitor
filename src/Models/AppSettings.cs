namespace ExcelTrainingMonitor.Models
{
    public class AppSettings
    {
        public string ExcelPath { get; set; }

        public int ScanIntervalMinutes { get; set; } = 5;

        public TimeSpan ReminderTime { get; set; }

        public List<string> ReminderDays { get; set; }

        public int StaleDaysThreshold { get; set; } = 7;
    }
}