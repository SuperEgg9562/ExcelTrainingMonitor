using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace ExcelTrainingMonitor.Controls
{
    internal sealed class PieChartPanel : Control
    {
        private readonly List<PieSegment> segments = new List<PieSegment>();

        public PieChartPanel()
        {
            DoubleBuffered = true;
            MinimumSize = new Size(220, 220);
            Font = new Font("Segoe UI", 10F);
        }

        public string ChartTitle { get; set; } = "";

        public void SetSegments(IEnumerable<PieSegment> values)
        {
            segments.Clear();
            segments.AddRange(values.Where(x => x.Value > 0));
            Invalidate();
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);

            e.Graphics.Clear(Color.FromArgb(8, 10, 9));
            e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;

            using var titleBrush = new SolidBrush(Color.FromArgb(0, 255, 40));
            using var textBrush = new SolidBrush(Color.FromArgb(120, 255, 135));
            using var borderPen = new Pen(Color.FromArgb(0, 220, 35));
            using var titleFont = new Font(Font, FontStyle.Bold);

            e.Graphics.DrawString(ChartTitle, titleFont, titleBrush, 12, 12);
            e.Graphics.DrawRectangle(borderPen, 0, 0, Width - 1, Height - 1);

            int total = segments.Sum(x => x.Value);
            if (total == 0)
            {
                e.Graphics.DrawString("No data", Font, textBrush, 12, 44);
                return;
            }

            int legendWidth = 180;
            int chartSize = Math.Min(Height - 70, Width - legendWidth - 40);
            chartSize = Math.Max(chartSize, 120);
            var chartRect = new Rectangle(16, 48, chartSize, chartSize);

            float startAngle = -90F;
            foreach (var segment in segments)
            {
                float sweepAngle = 360F * segment.Value / total;
                using var brush = new SolidBrush(segment.Color);
                e.Graphics.FillPie(brush, chartRect, startAngle, sweepAngle);
                startAngle += sweepAngle;
            }

            using var outlinePen = new Pen(Color.FromArgb(0, 220, 35), 2F);
            e.Graphics.DrawEllipse(outlinePen, chartRect);

            int legendX = chartRect.Right + 20;
            int legendY = 52;

            foreach (var segment in segments)
            {
                using var brush = new SolidBrush(segment.Color);
                e.Graphics.FillRectangle(brush, legendX, legendY + 3, 14, 14);
                e.Graphics.DrawString(
                    $"{segment.Label}: {segment.Value}",
                    Font,
                    textBrush,
                    legendX + 22,
                    legendY);
                legendY += 28;
            }
        }
    }

    internal sealed class PieSegment
    {
        public string Label { get; set; } = "";
        public int Value { get; set; }
        public Color Color { get; set; }
    }
}
