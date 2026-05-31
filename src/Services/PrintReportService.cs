using ExcelTrainingMonitor.Models;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Printing;
using System.Linq;
using System.Windows.Forms;

namespace ExcelTrainingMonitor.Services
{
    internal static class PrintReportService
    {
        public static void PrintReport(
            IWin32Window owner,
            IEnumerable<TrainingAlert> alerts,
            IEnumerable<HistoryEntry> history,
            Control statusChart,
            Control openChart)
        {
            List<TrainingAlert> alertRows = alerts.ToList();
            List<HistoryEntry> historyRows = history.ToList();
            int page = 0;

            using var document = new PrintDocument();
            document.DocumentName = "Excel Training Monitor Report";
            document.PrintPage += (sender, e) =>
            {
                page++;
                DrawPage(e.Graphics, e.MarginBounds, page, alertRows, historyRows, statusChart, openChart);
                e.HasMorePages = false;
            };

            using var dialog = new PrintDialog
            {
                Document = document,
                UseEXDialog = true
            };

            if (dialog.ShowDialog(owner) == DialogResult.OK)
            {
                document.Print();
            }
        }

        private static void DrawPage(
            Graphics graphics,
            Rectangle bounds,
            int page,
            List<TrainingAlert> alerts,
            List<HistoryEntry> history,
            Control statusChart,
            Control openChart)
        {
            using var titleFont = new Font("Segoe UI", 16F, FontStyle.Bold);
            using var headerFont = new Font("Segoe UI", 10F, FontStyle.Bold);
            using var font = new Font("Segoe UI", 9F);
            using var brush = new SolidBrush(Color.Black);
            using var linePen = new Pen(Color.Black);

            int y = bounds.Top;
            graphics.DrawString("Excel Training Monitor Report", titleFont, brush, bounds.Left, y);
            y += 34;
            graphics.DrawString(DateTime.Now.ToString("yyyy-MM-dd HH:mm"), font, brush, bounds.Left, y);
            y += 28;

            DrawSummary(graphics, bounds.Left, ref y, alerts, headerFont, font, brush);
            y += 12;

            DrawChart(graphics, statusChart, bounds.Left, y, bounds.Width / 2 - 12, 180);
            DrawChart(graphics, openChart, bounds.Left + bounds.Width / 2 + 12, y, bounds.Width / 2 - 12, 180);
            y += 200;

            DrawAlertsTable(graphics, bounds, ref y, alerts.Take(20).ToList(), headerFont, font, brush, linePen);
            y += 16;
            DrawHistoryTable(graphics, bounds, ref y, history.Take(12).ToList(), headerFont, font, brush);

            graphics.DrawString($"Page {page}", font, brush, bounds.Right - 60, bounds.Bottom + 20);
        }

        private static void DrawSummary(Graphics graphics, int x, ref int y, List<TrainingAlert> alerts, Font headerFont, Font font, Brush brush)
        {
            graphics.DrawString("Summary", headerFont, brush, x, y);
            y += 22;
            graphics.DrawString($"Total: {alerts.Count}", font, brush, x, y);
            graphics.DrawString($"Not Trained: {alerts.Count(a => a.Status == "Not Trained")}", font, brush, x + 120, y);
            graphics.DrawString($"In Training: {alerts.Count(a => a.Status == "In Training")}", font, brush, x + 280, y);
            graphics.DrawString($"Complete: {alerts.Count(a => a.Status == "Complete")}", font, brush, x + 440, y);
            y += 24;
        }

        private static void DrawChart(Graphics graphics, Control chart, int x, int y, int width, int height)
        {
            if (chart.Width <= 0 || chart.Height <= 0)
                return;

            using var bitmap = new Bitmap(chart.Width, chart.Height);
            chart.DrawToBitmap(bitmap, new Rectangle(Point.Empty, chart.Size));
            graphics.DrawImage(bitmap, new Rectangle(x, y, width, height));
        }

        private static void DrawAlertsTable(Graphics graphics, Rectangle bounds, ref int y, List<TrainingAlert> alerts, Font headerFont, Font font, Brush brush, Pen linePen)
        {
            graphics.DrawString("Current Training Grid", headerFont, brush, bounds.Left, y);
            y += 24;

            graphics.DrawLine(linePen, bounds.Left, y, bounds.Right, y);
            y += 6;
            graphics.DrawString("Employee", headerFont, brush, bounds.Left, y);
            graphics.DrawString("Category", headerFont, brush, bounds.Left + 190, y);
            graphics.DrawString("Status", headerFont, brush, bounds.Left + 390, y);
            y += 22;

            foreach (TrainingAlert alert in alerts)
            {
                graphics.DrawString(alert.EmployeeName ?? "", font, brush, bounds.Left, y);
                graphics.DrawString(alert.Category ?? "", font, brush, bounds.Left + 190, y);
                graphics.DrawString(alert.Status ?? "", font, brush, bounds.Left + 390, y);
                y += 20;
            }
        }

        private static void DrawHistoryTable(Graphics graphics, Rectangle bounds, ref int y, List<HistoryEntry> history, Font headerFont, Font font, Brush brush)
        {
            graphics.DrawString("Recent Updates", headerFont, brush, bounds.Left, y);
            y += 24;

            foreach (HistoryEntry entry in history)
            {
                graphics.DrawString(
                    $"{entry.Timestamp:yyyy-MM-dd HH:mm}  {entry.Employee}  {entry.Category}  {entry.OldStatus} -> {entry.NewStatus}",
                    font,
                    brush,
                    bounds.Left,
                    y);
                y += 20;
            }
        }
    }
}
