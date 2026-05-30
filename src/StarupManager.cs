using Microsoft.Win32;
using System.Windows.Forms;

namespace ExcelTrainingMonitor
{
    public static class StartupManager
    {
        public static void EnableStartup()
        {
            RegistryKey key =
                Registry.CurrentUser.OpenSubKey(
                    @"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run",
                    true);

            key.SetValue(
                "ExcelTrainingMonitor",
                Application.ExecutablePath);
        }

        public static void DisableStartup()
        {
            RegistryKey key =
                Registry.CurrentUser.OpenSubKey(
                    @"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run",
                    true);

            key.DeleteValue(
                "ExcelTrainingMonitor",
                false);
        }
    }
}