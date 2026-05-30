using System;
using System.Windows.Forms;

namespace ExcelTrainingMonitor
{
    internal static class Program
    {
        [STAThread]
        static void Main()
        {
            ApplicationConfiguration.Initialize();

            Application.Run(new MainForm());
        }
    }
}