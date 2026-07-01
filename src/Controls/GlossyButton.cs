using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace ExcelTrainingMonitor.Controls
{
    internal sealed class GlossyButton : Button
    {
        private readonly PointerInteractionState pointerState = new PointerInteractionState();
        public Color AccentColor { get; set; } = Color.FromArgb(0, 255, 40);
        public Color HoverBackColor { get; set; } = Color.FromArgb(7, 92, 24);

        public GlossyButton()
        {
            FlatStyle = FlatStyle.Flat;
            FlatAppearance.BorderSize = 0;
            UseVisualStyleBackColor = false;
            DoubleBuffered = true;
            ForeColor = Color.FromArgb(0, 255, 40);
            BackColor = Color.FromArgb(3, 45, 13);
            pointerState.Attach(this);
        }

        protected override void OnPaint(PaintEventArgs pevent)
        {
            Graphics g = pevent.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;

            Rectangle rect = new Rectangle(0, 0, Width - 1, Height - 1);
            Rectangle inner = new Rectangle(1, 1, Width - 3, Height - 3);
            Rectangle gloss = new Rectangle(2, 2, Width - 5, Math.Max(8, Height / 2));

            Color top = pointerState.Pressing
                ? ControlPaint.Dark(BackColor, 0.35F)
                : pointerState.Hovering ? ControlPaint.Light(HoverBackColor, 0.12F) : ControlPaint.Light(BackColor, 0.12F);
            Color bottom = pointerState.Pressing
                ? ControlPaint.Dark(BackColor, 0.12F)
                : pointerState.Hovering ? HoverBackColor : ControlPaint.Dark(BackColor, 0.18F);

            using var bodyBrush = new LinearGradientBrush(inner, top, bottom, LinearGradientMode.Vertical);
            using var glossBrush = new LinearGradientBrush(
                gloss,
                Color.FromArgb(75, AccentColor),
                Color.FromArgb(8, AccentColor),
                LinearGradientMode.Vertical);
            using var borderPen = new Pen(AccentColor);
            using var darkPen = new Pen(ControlPaint.Dark(BackColor, 0.6F));

            g.FillRectangle(bodyBrush, inner);
            g.FillRectangle(glossBrush, gloss);
            g.DrawRectangle(darkPen, rect);
            g.DrawRectangle(borderPen, 1, 1, Width - 3, Height - 3);
            g.DrawLine(Pens.Black, 2, Height - 2, Width - 3, Height - 2);

            TextRenderer.DrawText(
                g,
                Text,
                Font,
                ClientRectangle,
                ForeColor,
                TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis);
        }
    }
}
