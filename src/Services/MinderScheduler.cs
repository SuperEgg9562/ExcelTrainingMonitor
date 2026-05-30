using System;
using System.IO;

namespace ExcelTrainingMonitor
{
    public static class MinderScheduler
    {
        public static bool FileNeedsReminder(string path, int hoursWithoutUpdate)
        {
            if (!File.Exists(path))
                return false;

            DateTime lastWrite = File.GetLastWriteTime(path);

            return (DateTime.Now - lastWrite).TotalHours >= hoursWithoutUpdate;
        }
    }
}