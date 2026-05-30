using System;

namespace ExcelTrainingMonitor.Models
{
    public class ReminderRule
    {
        public DayOfWeek Day { get; set; }
        public TimeSpan Time { get; set; }
    }
}