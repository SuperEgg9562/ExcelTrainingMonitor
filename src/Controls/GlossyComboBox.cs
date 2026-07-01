using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace ExcelTrainingMonitor.Controls
{
    internal sealed class GlossyComboBox : ComboBox
    {
        public Color AccentColor { get; set; } = Color.FromArgb(0, 255, 40);
        public Color HoverBackColor { get; set; } = Color.FromArgb(18, 96, 18);

        public GlossyComboBox()
        {
            DrawMode = DrawMode.OwnerDrawFixed;
            DropDownStyle = ComboBoxStyle.DropDownList;
            FlatStyle = FlatStyle.Flat;
            ForeColor = Color.FromArgb(0, 255, 40);
            BackColor = Color.FromArgb(2, 34, 10);
            ItemHeight = 24;
        }

        protected override void OnDrawItem(DrawItemEventArgs e)
        {
            if (e.Index < 0)
                return;

            bool selected = (e.State & DrawItemState.Selected) == DrawItemState.Selected;
            Rectangle bounds = e.Bounds;
            Color top = selected ? ControlPaint.Light(HoverBackColor, 0.12F) : ControlPaint.Light(BackColor, 0.1F);
            Color bottom = selected ? HoverBackColor : ControlPaint.Dark(BackColor, 0.22F);
            Color text = ForeColor;

            using var bodyBrush = new LinearGradientBrush(bounds, top, bottom, LinearGradientMode.Vertical);
            using var glossBrush = new LinearGradientBrush(
                new Rectangle(bounds.X, bounds.Y, bounds.Width, Math.Max(4, bounds.Height / 2)),
                Color.FromArgb(selected ? 70 : 30, AccentColor),
                Color.FromArgb(0, AccentColor),
                LinearGradientMode.Vertical);

            e.Graphics.FillRectangle(bodyBrush, bounds);
            e.Graphics.FillRectangle(glossBrush, bounds.X, bounds.Y, bounds.Width, Math.Max(4, bounds.Height / 2));

            TextRenderer.DrawText(
                e.Graphics,
                GetItemText(Items[e.Index]),
                Font,
                new Rectangle(bounds.X + 8, bounds.Y, bounds.Width - 12, bounds.Height),
                text,
                TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis);
        }

        protected override void WndProc(ref Message m)
        {
            base.WndProc(ref m);

            const int wmPaint = 0x000F;
            const int wmNcPaint = 0x0085;

            if (m.Msg == wmPaint || m.Msg == wmNcPaint)
            {
                using Graphics g = CreateGraphics();
                Rectangle border = new Rectangle(0, 0, Width - 1, Height - 1);
                Rectangle inner = Rectangle.Inflate(border, -1, -1);
                Rectangle arrowBox = new Rectangle(Math.Max(0, Width - 24), 1, 23, Math.Max(1, Height - 2));
                Rectangle textBox = new Rectangle(8, 1, Math.Max(1, Width - 34), Math.Max(1, Height - 2));

                using var borderPen = new Pen(AccentColor);
                using var darkPen = new Pen(Color.Black);
                using var bodyBrush = new LinearGradientBrush(
                    inner,
                    ControlPaint.Light(BackColor, 0.12F),
                    ControlPaint.Dark(BackColor, 0.22F),
                    LinearGradientMode.Vertical);
                using var shineBrush = new LinearGradientBrush(
                    new Rectangle(inner.X, inner.Y, inner.Width, Math.Max(5, inner.Height / 2)),
                    Color.FromArgb(45, AccentColor),
                    Color.FromArgb(0, AccentColor),
                    LinearGradientMode.Vertical);
                using var arrowBrush = new LinearGradientBrush(
                    arrowBox,
                    ControlPaint.Light(HoverBackColor, 0.08F),
                    ControlPaint.Dark(HoverBackColor, 0.25F),
                    LinearGradientMode.Vertical);
                using var arrowPen = new Pen(AccentColor, 2F);

                g.FillRectangle(bodyBrush, inner);
                g.FillRectangle(shineBrush, inner.X, inner.Y, inner.Width, Math.Max(5, inner.Height / 2));
                g.FillRectangle(arrowBrush, arrowBox);
                g.DrawRectangle(darkPen, border);
                g.DrawRectangle(borderPen, inner);
                g.DrawLine(borderPen, arrowBox.X, arrowBox.Y + 2, arrowBox.X, arrowBox.Bottom - 3);

                int midX = arrowBox.X + arrowBox.Width / 2;
                int midY = arrowBox.Y + arrowBox.Height / 2 + 1;
                g.DrawLines(
                    arrowPen,
                    new[]
                    {
                        new Point(midX - 5, midY - 2),
                        new Point(midX, midY + 3),
                        new Point(midX + 5, midY - 2)
                    });

                TextRenderer.DrawText(
                    g,
                    Text,
                    Font,
                    textBox,
                    ForeColor,
                    TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis);
            }
        }
    }
}
