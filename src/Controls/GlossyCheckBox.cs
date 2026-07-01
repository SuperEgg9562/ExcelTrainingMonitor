using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace ExcelTrainingMonitor.Controls
{
    internal sealed class GlossyCheckBox : CheckBox
    {
        private readonly PointerInteractionState pointerState = new PointerInteractionState();
        public Color AccentColor { get; set; } = Color.FromArgb(0, 255, 40);
        public Color BoxBackColor { get; set; } = Color.FromArgb(12, 62, 12);
        public Color HoverBackColor { get; set; } = Color.FromArgb(28, 96, 14);

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
                ? ControlPaint.Dark(BoxBackColor, 0.3F)
                : pointerState.Hovering ? HoverBackColor : ControlPaint.Light(BoxBackColor, 0.1F);
            Color bottom = Checked
                ? ControlPaint.Dark(AccentColor, 0.25F)
                : ControlPaint.Dark(BoxBackColor, 0.25F);

            using var bodyBrush = new LinearGradientBrush(inner, top, bottom, LinearGradientMode.Vertical);
            using var glossBrush = new LinearGradientBrush(
                gloss,
                Color.FromArgb(65, AccentColor),
                Color.FromArgb(8, AccentColor),
                LinearGradientMode.Vertical);
            using var borderPen = new Pen(AccentColor);
            using var shadowPen = new Pen(Color.Black);

            g.FillRectangle(bodyBrush, inner);
            g.FillRectangle(glossBrush, gloss);
            g.DrawRectangle(shadowPen, box);
            g.DrawRectangle(borderPen, box.X + 1, box.Y + 1, box.Width - 2, box.Height - 2);

            if (Checked)
            {
                using var checkPen = new Pen(AccentColor, 2.2F)
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
