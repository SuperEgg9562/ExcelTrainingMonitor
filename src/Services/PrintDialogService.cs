using System.Drawing.Printing;
using System.Windows.Forms;

namespace ExcelTrainingMonitor.Services
{
    internal static class PrintDialogService
    {
        public static void Print(IWin32Window owner, PrintDocument document)
        {
            using var dialog = new PrintDialog
            {
                Document = document,
                UseEXDialog = true
            };

            if (dialog.ShowDialog(owner) == DialogResult.OK)
                document.Print();
        }
    }
}
