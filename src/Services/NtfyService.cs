using ExcelTrainingMonitor.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTrainingMonitor.Services
{
    internal static class NtfyService
    {
        private static readonly HttpClient Client = new HttpClient();

        public static async Task SendReminderAsync(AppSettings settings, IEnumerable<TrainingAlert> alerts)
        {
            if (!settings.NtfyEnabled || string.IsNullOrWhiteSpace(settings.NtfyTopic))
                return;

            List<TrainingAlert> allOpenAlerts = alerts
                .Where(x => x.Status == "Not Trained" || x.Status == "In Training")
                .OrderBy(x => x.EmployeeName)
                .ThenBy(x => x.Category)
                .ToList();
            List<TrainingAlert> openAlerts = allOpenAlerts
                .Take(12)
                .ToList();

            var message = new StringBuilder();
            message.AppendLine($"{allOpenAlerts.Count} open training item(s)");

            foreach (TrainingAlert alert in openAlerts)
            {
                message.AppendLine($"{alert.EmployeeName}: {alert.Category} ({alert.Status})");
            }

            string topic = Uri.EscapeDataString(settings.NtfyTopic.Trim());
            using var request = new HttpRequestMessage(HttpMethod.Post, $"https://ntfy.sh/{topic}")
            {
                Content = new StringContent(message.ToString().Trim(), Encoding.UTF8, "text/plain")
            };

            request.Headers.TryAddWithoutValidation("Title", "Training Reminder");
            request.Headers.TryAddWithoutValidation("Tags", "warning,calendar");
            request.Headers.TryAddWithoutValidation("Priority", "3");

            if (!string.IsNullOrWhiteSpace(settings.NtfyEmail))
            {
                request.Headers.TryAddWithoutValidation("Email", settings.NtfyEmail.Trim());
            }

            using HttpResponseMessage response = await Client.SendAsync(request);
            response.EnsureSuccessStatusCode();
        }
    }
}
