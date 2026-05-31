using System.Windows.Forms;

namespace ExcelTrainingMonitor.Services
{
    public static class NotificationService
    {
        public static void Show(string title, string message)
        {
            MessageBox.Show(message, title, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
