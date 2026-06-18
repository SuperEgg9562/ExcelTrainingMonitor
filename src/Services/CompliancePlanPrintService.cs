using System;
using System.Data;
using System.Drawing;
using System.Drawing.Printing;
using System.Linq;
using System.Windows.Forms;

namespace ExcelTrainingMonitor.Services
{
    internal static class CompliancePlanPrintService
    {
        public static void PrintProcessRecord(
            IWin32Window owner,
            string version,
            string title,
            DateTime processDateTime,
            DataTable table,
            Image logo)
        {
            Print(owner, version, title, processDateTime, "", "", table, logo);
        }

        public static void Print(
            IWin32Window owner,
            string technicalTerms,
            string title,
            DateTime planDateTime,
            string legend,
            string detailIssues,
            DataTable table,
            Image logo)
        {
            int nextRow = 0;
            int pageNumber = 0;

            using var document = new PrintDocument
            {
                DocumentName = string.IsNullOrWhiteSpace(title) ? "Compliance Plan" : title
            };
            document.DefaultPageSettings.Landscape = true;
            document.PrintPage += (sender, e) =>
            {
                pageNumber++;
                DrawPage(
                    e.Graphics,
                    e.MarginBounds,
                    technicalTerms,
                    title,
                    planDateTime,
                    legend,
                    detailIssues,
                    table,
                    logo,
                    pageNumber,
                    ref nextRow,
                    out bool hasMorePages);
                e.HasMorePages = hasMorePages;
            };

            PrintDialogService.Print(owner, document);
        }

        private static void DrawPage(
            Graphics graphics,
            Rectangle bounds,
            string technicalTerms,
            string title,
            DateTime planDateTime,
            string legend,
            string detailIssues,
            DataTable table,
            Image logo,
            int pageNumber,
            ref int nextRow,
            out bool hasMorePages)
        {
            using var titleFont = new Font("Segoe UI", 16F, FontStyle.Bold);
            using var headingFont = new Font("Segoe UI", 9F, FontStyle.Bold);
            using var textFont = new Font("Segoe UI", 8F);
            using var brush = new SolidBrush(Color.Black);
            using var gridPen = new Pen(Color.Black, 0.8F);
            using var centered = new StringFormat
            {
                Alignment = StringAlignment.Center,
                LineAlignment = StringAlignment.Center,
                Trimming = StringTrimming.EllipsisCharacter
            };

            float y = bounds.Top;
            if (pageNumber == 1)
            {
                y = DrawTextSection(graphics, technicalTerms, textFont, brush, bounds.Left, y, bounds.Width, 72F);
                y += 8F;

                float logoWidth = logo == null ? 0F : 100F;
                var titleBounds = new RectangleF(bounds.Left, y, bounds.Width - logoWidth - 8F, 70F);
                graphics.DrawString(string.IsNullOrWhiteSpace(title) ? "Compliance Plan" : title, titleFont, brush, titleBounds, centered);
                graphics.DrawRectangle(gridPen, Rectangle.Round(titleBounds));

                if (logo != null)
                {
                    var logoBounds = new RectangleF(bounds.Right - logoWidth, y, logoWidth, 70F);
                    graphics.DrawImage(logo, logoBounds);
                    graphics.DrawRectangle(gridPen, Rectangle.Round(logoBounds));
                }

                y += 80F;
                graphics.DrawString($"Date and time: {planDateTime:yyyy-MM-dd HH:mm}", headingFont, brush, bounds.Left, y);
                y += 24F;
                y = DrawTextSection(graphics, legend, textFont, brush, bounds.Left, y, bounds.Width, 72F);
                y += 10F;
            }
            else
            {
                graphics.DrawString($"{title} — continued", headingFont, brush, bounds.Left, y);
                graphics.DrawString($"Page {pageNumber}", textFont, brush, bounds.Right - 55F, y);
                y += 26F;
            }

            int columnCount = Math.Max(1, table.Columns.Count);
            float columnWidth = bounds.Width / (float)columnCount;
            const float headerHeight = 26F;
            DrawGridRow(graphics, table.Columns.Cast<DataColumn>().Select(column => column.ColumnName).ToArray(),
                headingFont, brush, gridPen, centered, bounds.Left, y, columnWidth, headerHeight);
            y += headerHeight;

            float printableBottom = bounds.Bottom - 170F;
            while (nextRow < table.Rows.Count)
            {
                string[] values = table.Rows[nextRow].ItemArray.Select(value => value?.ToString() ?? "").ToArray();
                float rowHeight = MeasureRowHeight(graphics, values, textFont, columnWidth);
                if (y + rowHeight > printableBottom)
                    break;

                DrawGridRow(graphics, values, textFont, brush, gridPen, centered, bounds.Left, y, columnWidth, rowHeight);
                y += rowHeight;
                nextRow++;
            }

            hasMorePages = nextRow < table.Rows.Count;
            if (!hasMorePages)
            {
                float signY = bounds.Bottom - 154F;
                graphics.DrawString("Completed by: __________", headingFont, brush, bounds.Left, signY);
                graphics.DrawString("Signature: __________", headingFont, brush, bounds.Left + 260F, signY);

                float detailTitleY = signY + 28F;
                graphics.DrawString("Detail issues with corrective actions", headingFont, brush, bounds.Left, detailTitleY);
                var detailBounds = new RectangleF(bounds.Left, detailTitleY + 20F, bounds.Width, 100F);
                graphics.DrawRectangle(gridPen, Rectangle.Round(detailBounds));
                detailBounds.Inflate(-5F, -5F);
                graphics.DrawString(detailIssues ?? "", textFont, brush, detailBounds);
            }
        }

        private static float DrawTextSection(Graphics graphics, string text, Font font, Brush brush, float x, float y, float width, float maximumHeight)
        {
            if (string.IsNullOrWhiteSpace(text))
                return y;

            SizeF measured = graphics.MeasureString(text, font, new SizeF(width, maximumHeight));
            float height = Math.Min(maximumHeight, Math.Max(font.Height + 8F, measured.Height + 8F));
            var rectangle = new RectangleF(x, y, width, height);
            graphics.DrawRectangle(Pens.Black, Rectangle.Round(rectangle));
            rectangle.Inflate(-4F, -4F);
            graphics.DrawString(text, font, brush, rectangle);
            return y + height;
        }

        private static float MeasureRowHeight(Graphics graphics, string[] values, Font font, float columnWidth)
        {
            float contentWidth = Math.Max(10F, columnWidth - 8F);
            float required = values
                .Select(value => graphics.MeasureString(string.IsNullOrEmpty(value) ? " " : value, font, new SizeF(contentWidth, 90F)).Height + 8F)
                .DefaultIfEmpty(26F)
                .Max();
            return Math.Clamp(required, 26F, 90F);
        }

        private static void DrawGridRow(
            Graphics graphics,
            string[] values,
            Font font,
            Brush brush,
            Pen pen,
            StringFormat format,
            float x,
            float y,
            float columnWidth,
            float rowHeight)
        {
            for (int column = 0; column < values.Length; column++)
            {
                var cell = new RectangleF(x + column * columnWidth, y, columnWidth, rowHeight);
                graphics.DrawRectangle(pen, Rectangle.Round(cell));
                cell.Inflate(-3F, -2F);
                graphics.DrawString(values[column], font, brush, cell, format);
            }
        }
    }
}
