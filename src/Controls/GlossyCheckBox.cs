using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace ExcelTrainingMonitor.Controls
{
    internal sealed class GlossyCheckBox : CheckBox
    {
        private readonly PointerInteractionState pointerState = new PointerInteractionState();

        public GlossyCheckBox()
        {
            AutoSize = true;
            DoubleBuffered = true;
            FlatStyle = FlatStyle.Flat;
            ForeColor = Color.FromArgb(0, 255, 40);
            BackColor = Color.Transparent;
            UseVisualStyleBackColor = false;
            pointerState.Attach(this);
        }

        protected override void OnCheckedChanged(EventArgs e)
        {
            Invalidate();
            base.OnCheckedChanged(e);
        }

        protected override void OnPaint(PaintEventArgs pevent)
        {
            Graphics g = pevent.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.Clear(Parent?.BackColor ?? Color.Black);

            int boxSize = Math.Min(18, Math.Max(14, Height - 5));
            var box = new Rectangle(1, (Height - boxSize) / 2, boxSize, boxSize);
            var inner = Rectangle.Inflate(box, -2, -2);
            var gloss = new Rectangle(inner.X, inner.Y, inner.Width, Math.Max(5, inner.Height / 2));

            Color top = pointerState.Pressing
                ? Color.FromArgb(18, 50, 8)
                : pointerState.Hovering ? Color.FromArgb(28, 96, 14) : Color.FromArgb(12, 62, 12);
            Color bottom = Checked
                ? Color.FromArgb(0, 150, 30)
                : Color.FromArgb(0, 28, 8);

            using var bodyBrush = new LinearGradientBrush(inner, top, bottom, LinearGradientMode.Vertical);
            using var glossBrush = new LinearGradientBrush(
                gloss,
                Color.FromArgb(145, 255, 236, 125),
                Color.FromArgb(20, 255, 190, 20),
                LinearGradientMode.Vertical);
            using var borderPen = new Pen(pointerState.Hovering ? Color.FromArgb(255, 215, 35) : Color.FromArgb(0, 255, 40));
            using var shadowPen = new Pen(Color.Black);

            g.FillRectangle(bodyBrush, inner);
            g.FillRectangle(glossBrush, gloss);
            g.DrawRectangle(shadowPen, box);
            g.DrawRectangle(borderPen, box.X + 1, box.Y + 1, box.Width - 2, box.Height - 2);

            if (Checked)
            {
                using var checkPen = new Pen(Color.FromArgb(255, 220, 20), 2.2F)
                {
                    StartCap = LineCap.Round,
                    EndCap = LineCap.Round
                };

                Point p1 = new Point(box.X + 4, box.Y + box.Height / 2);
                Point p2 = new Point(box.X + box.Width / 2 - 1, box.Bottom - 5);
                Point p3 = new Point(box.Right - 4, box.Y + 4);
                g.DrawLines(checkPen, new[] { p1, p2, p3 });
            }

            Rectangle textBounds = new Rectangle(box.Right + 7, 0, Width - box.Right - 7, Height);
            TextRenderer.DrawText(
                g,
                Text,
                Font,
                textBounds,
                ForeColor,
                TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis);
        }
    }
}
