using ExcelTrainingMonitor.Models;
using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelTrainingMonitor.Services
{
    internal sealed class ReminderAgentContext : ApplicationContext
    {
        private readonly NotifyIcon notifyIcon;
        private readonly System.Windows.Forms.Timer timer;
        private bool isChecking;

        public ReminderAgentContext()
        {
            notifyIcon = new NotifyIcon
            {
                Text = "Excel Training Reminder Agent",
                Visible = true,
                Icon = SystemIcons.Information,
                ContextMenuStrip = new ContextMenuStrip()
            };
            notifyIcon.ContextMenuStrip.Items.Add("Exit Reminder Agent", null, (s, e) => ExitThread());

            timer = new System.Windows.Forms.Timer { Interval = 30_000 };
            timer.Tick += async (s, e) => await CheckReminderAsync();
            timer.Start();

            _ = CheckReminderAsync();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                timer.Stop();
                timer.Dispose();
                notifyIcon.Visible = false;
                notifyIcon.Dispose();
            }

            base.Dispose(disposing);
        }

        private async Task CheckReminderAsync()
        {
            if (isChecking)
                return;

            isChecking = true;
            try
            {
                AppSettings settings = SettingsManager.Load();
                if (!settings.ReminderEnabled)
                    return;

                DateTime reminderDate = settings.ReminderDateTime;
                if (DateTime.Now < reminderDate || settings.LastAgentReminderSentFor == reminderDate)
                    return;

                var alerts = File.Exists(settings.ExcelPath)
                    ? ExcelMonitor.ScanFile(settings.ExcelPath)
                    : new System.Collections.Generic.List<TrainingAlert>();

                var open = alerts
                    .Where(x => x.Status == "Not Trained" || x.Status == "In Training")
                    .Take(5)
                    .ToList();

                string message = open.Count == 0
                    ? $"Reminder due: {reminderDate:yyyy-MM-dd HH:mm}"
                    : string.Join(Environment.NewLine, open.Select(x => $"{x.EmployeeName}: {x.Category} ({x.Status})"));

                notifyIcon.BalloonTipTitle = "Training Reminder";
                notifyIcon.BalloonTipText = message.Length > 250 ? message[..250] : message;
                notifyIcon.ShowBalloonTip(8000);

                settings.LastAgentReminderSentFor = reminderDate;
                SettingsManager.Save(settings);

                await NtfyService.SendReminderAsync(settings, alerts);
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"Reminder agent check failed: {ex}");
            }
            finally
            {
                isChecking = false;
            }
        }
    }
}
