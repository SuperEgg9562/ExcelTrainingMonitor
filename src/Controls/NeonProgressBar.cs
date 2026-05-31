using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace ExcelTrainingMonitor.Controls
{
    internal sealed class NeonProgressBar : Control
    {
        private int maximum = 100;
        private int value;

        public NeonProgressBar()
        {
            DoubleBuffered = true;
            MinimumSize = new Size(80, 18);
        }

        public int Maximum
        {
            get => maximum;
            set
            {
                maximum = Math.Max(1, value);
                this.value = Math.Min(this.value, maximum);
                Invalidate();
            }
        }

        public int Value
        {
            get => value;
            set
            {
                this.value = Math.Max(0, Math.Min(Maximum, value));
                Invalidate();
            }
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);

            var accent = Color.FromArgb(0, 255, 40);
            var fillTop = Color.FromArgb(110, 255, 115);
            var fillMid = Color.FromArgb(0, 180, 35);
            var fillBottom = Color.FromArgb(0, 70, 15);
            var back = Color.FromArgb(2, 16, 6);

            e.Graphics.Clear(back);

            Rectangle border = new Rectangle(0, 0, Width - 1, Height - 1);
            Rectangle inner = new Rectangle(2, 2, Math.Max(0, Width - 4), Math.Max(0, Height - 4));
            int fillWidth = Maximum <= 0 ? 0 : inner.Width * Value / Maximum;

            using var borderPen = new Pen(accent);
            using var glowPen = new Pen(Color.FromArgb(90, 0, 255, 60));
            using var backBrush = new LinearGradientBrush(inner, Color.FromArgb(15, 15, 15), back, LinearGradientMode.Vertical);

            e.Graphics.FillRectangle(backBrush, inner);

            if (fillWidth > 0)
            {
                var fillRect = new Rectangle(inner.X, inner.Y, fillWidth, inner.Height);
                using var fillBrush = new LinearGradientBrush(fillRect, fillTop, fillBottom, LinearGradientMode.Vertical);
                using var shineBrush = new LinearGradientBrush(
                    new Rectangle(fillRect.X, fillRect.Y, fillRect.Width, Math.Max(4, fillRect.Height / 2)),
                    Color.FromArgb(180, 220, 255, 220),
                    Color.FromArgb(20, 0, 255, 40),
                    LinearGradientMode.Vertical);

                e.Graphics.FillRectangle(fillBrush, fillRect);
                e.Graphics.FillRectangle(shineBrush, fillRect.X, fillRect.Y, fillRect.Width, Math.Max(4, fillRect.Height / 2));
            }

            e.Graphics.DrawRectangle(glowPen, 1, 1, Width - 3, Height - 3);
            e.Graphics.DrawRectangle(borderPen, border);
        }
    }
}
