using System.Drawing;
using System.Drawing.Imaging;
using System.Windows.Forms;

namespace ExcelTrainingMonitor.Services
{
    internal static class ChartExportService
    {
        public static void Export(Control control, string path)
        {
            using var bitmap = new Bitmap(control.Width, control.Height);
            control.DrawToBitmap(bitmap, new Rectangle(Point.Empty, control.Size));
            bitmap.Save(path, ImageFormat.Png);
        }
    }
}
