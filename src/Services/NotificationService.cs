using System.Windows.Forms;

namespace ExcelTrainingMonitor.Services
{
    public static class NotificationService
    {
        public static void Show(string title, string message)
        {
            // For portability and to avoid extra runtime requirements, use MessageBox for now.
            MessageBox.Show(message, title, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
