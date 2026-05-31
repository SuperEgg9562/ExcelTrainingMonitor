using System.IO;
using System.Text.Json;
using ExcelTrainingMonitor.Models;

namespace ExcelTrainingMonitor.Services
{
    public static class SettingsManager
    {
        private static string AppDataPath =>
            Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "ExcelTrainingMonitor");

        private static string ConfigPath => Path.Combine(AppDataPath, "config.txt");
        private static string SettingsPath => Path.Combine(AppDataPath, "settings.json");

        public static AppSettings Load()
        {
            try
            {
                if (File.Exists(SettingsPath))
                {
                    string json = File.ReadAllText(SettingsPath);
                    return JsonSerializer.Deserialize<AppSettings>(json) ?? new AppSettings();
                }
            }
            catch
            {
            }

            return new AppSettings
            {
                ExcelPath = LoadExcelPath()
            };
        }

        public static void Save(AppSettings settings)
        {
            Directory.CreateDirectory(AppDataPath);

            string json = JsonSerializer.Serialize(
                settings,
                new JsonSerializerOptions { WriteIndented = true });

            File.WriteAllText(SettingsPath, json);
        }

        public static void SaveExcelPath(string path)
        {
            AppSettings settings = Load();
            settings.ExcelPath = path;
            Save(settings);
        }

        public static string LoadExcelPath()
        {
            if (!File.Exists(ConfigPath))
                return "";

            return File.ReadAllText(ConfigPath);
        }
    }
}
