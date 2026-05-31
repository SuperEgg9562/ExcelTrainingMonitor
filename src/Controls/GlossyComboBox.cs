using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace ExcelTrainingMonitor.Controls
{
    internal sealed class GlossyComboBox : ComboBox
    {
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
            Color top = selected ? Color.FromArgb(18, 96, 18) : Color.FromArgb(5, 48, 12);
            Color bottom = selected ? Color.FromArgb(0, 140, 28) : Color.FromArgb(0, 18, 6);
            Color text = selected ? Color.FromArgb(255, 220, 20) : ForeColor;

            using var bodyBrush = new LinearGradientBrush(bounds, top, bottom, LinearGradientMode.Vertical);
            using var glossBrush = new LinearGradientBrush(
                new Rectangle(bounds.X, bounds.Y, bounds.Width, Math.Max(4, bounds.Height / 2)),
                Color.FromArgb(selected ? 120 : 80, 255, 236, 125),
                Color.FromArgb(0, 0, 255, 40),
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

                using var borderPen = new Pen(Color.FromArgb(0, 255, 40));
                using var darkPen = new Pen(Color.Black);
                using var bodyBrush = new LinearGradientBrush(
                    inner,
                    Color.FromArgb(8, 64, 14),
                    Color.FromArgb(0, 25, 7),
                    LinearGradientMode.Vertical);
                using var shineBrush = new LinearGradientBrush(
                    new Rectangle(inner.X, inner.Y, inner.Width, Math.Max(5, inner.Height / 2)),
                    Color.FromArgb(95, 255, 236, 125),
                    Color.FromArgb(0, 0, 255, 40),
                    LinearGradientMode.Vertical);
                using var arrowBrush = new LinearGradientBrush(
                    arrowBox,
                    Color.FromArgb(18, 95, 18),
                    Color.FromArgb(0, 35, 8),
                    LinearGradientMode.Vertical);
                using var arrowPen = new Pen(Color.FromArgb(255, 215, 35), 2F);

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
