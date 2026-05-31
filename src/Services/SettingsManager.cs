using System.IO;

namespace ExcelTrainingMonitor.Services
{
    public static class SettingsManager
    {
        private static string AppDataPath =>
            Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "ExcelTrainingMonitor");

        private static string ConfigPath => Path.Combine(AppDataPath, "config.txt");

        public static void SaveExcelPath(string path)
        {
            Directory.CreateDirectory(AppDataPath);
            File.WriteAllText(ConfigPath, path);
        }

        public static string LoadExcelPath()
        {
            if (!File.Exists(ConfigPath))
                return "";

            return File.ReadAllText(ConfigPath);
        }
    }
}
