using System;
using System.Threading;
using System.Windows.Forms;
using ExcelTrainingMonitor.Services;

namespace ExcelTrainingMonitor
{
    internal static class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            ApplicationConfiguration.Initialize();

            if (args.Length > 0 && args[0].Equals("--reminder-agent", StringComparison.OrdinalIgnoreCase))
            {
                using var mutex = new Mutex(true, "ExcelTrainingMonitor.ReminderAgent", out bool ownsMutex);
                if (!ownsMutex)
                    return;

                Application.Run(new ReminderAgentContext());
                return;
            }

            Application.Run(new MainForm());
        }
    }
}
