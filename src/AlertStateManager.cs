using ExcelTrainingMonitor.Models;
using System.Collections.Generic;
using System.Linq;

namespace ExcelTrainingMonitor
{
    public static class AlertStateManager
    {
        private static List<TrainingAlert> previousAlerts =
            new List<TrainingAlert>();

        public static List<TrainingAlert> GetNewAlerts(
            List<TrainingAlert> currentAlerts)
        {
            List<TrainingAlert> newAlerts =
                new List<TrainingAlert>();

            foreach (var current in currentAlerts)
            {
                bool exists =
                    previousAlerts.Any(p =>
                        p.EmployeeName == current.EmployeeName &&
                        p.Category == current.Category &&
                        p.Status == current.Status);

                if (!exists)
                {
                    newAlerts.Add(current);
                }
            }

            previousAlerts =
                currentAlerts
                .Select(a => new TrainingAlert
                {
                    EmployeeName = a.EmployeeName,
                    Category = a.Category,
                    Status = a.Status,
                    Timestamp = a.Timestamp
                })
                .ToList();

            return newAlerts;
        }
    }
}