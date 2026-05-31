using System.Windows.Forms;

namespace ExcelTrainingMonitor.Services
{
    public static class NotificationService
    {
        public static void Show(string title, string message)
        {
            using var form = new ScrollableNotificationForm(title, message);
            form.ShowDialog();
        }
    }
}
