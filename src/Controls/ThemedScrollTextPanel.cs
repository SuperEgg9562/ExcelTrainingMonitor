using ExcelTrainingMonitor.Services;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace ExcelTrainingMonitor.Controls
{
    internal sealed class ThemedScrollTextPanel : Control
    {
        private const int ScrollbarWidth = 14;
        private readonly List<string> wrappedLines = new List<string>();
        private AppTheme theme;
        private string text = "";
        private int firstLine;
        private bool draggingThumb;
        private int dragStartY;
        private int dragStartFirstLine;

        public ThemedScrollTextPanel()
        {
            DoubleBuffered = true;
            Font = new Font("Segoe UI", 10F);
            SetStyle(ControlStyles.Selectable, true);
        }

        public void SetTheme(AppTheme value)
        {
            theme = value;
            BackColor = theme.ControlBack;
            ForeColor = theme.Fore;
            Invalidate();
        }

        public void SetText(string value)
        {
            text = value ?? "";
            RebuildLines();
            Invalidate();
        }

        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
            RebuildLines();
        }

        protected override void OnMouseWheel(MouseEventArgs e)
        {
            base.OnMouseWheel(e);
            ScrollBy(e.Delta < 0 ? 3 : -3);
        }

        protected override void OnMouseDown(MouseEventArgs e)
        {
            Focus();
            Rectangle thumb = GetThumbRectangle();

            if (thumb.Contains(e.Location))
            {
                draggingThumb = true;
                dragStartY = e.Y;
                dragStartFirstLine = firstLine;
                Capture = true;
                return;
            }

            if (GetScrollbarTrackRectangle().Contains(e.Location))
            {
                ScrollBy(e.Y < thumb.Y ? -GetVisibleLineCount() : GetVisibleLineCount());
            }
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            if (!draggingThumb)
                return;

            int maxFirstLine = GetMaxFirstLine();
            Rectangle track = GetScrollbarTrackRectangle();
            Rectangle thumb = GetThumbRectangle();
            int movement = e.Y - dragStartY;
            int movablePixels = Math.Max(track.Height - thumb.Height, 1);
            int lineDelta = (int)Math.Round((double)movement / movablePixels * maxFirstLine);

            firstLine = Math.Max(0, Math.Min(maxFirstLine, dragStartFirstLine + lineDelta));
            Invalidate();
        }

        protected override void OnMouseUp(MouseEventArgs e)
        {
            draggingThumb = false;
            Capture = false;
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);

            AppTheme activeTheme = theme ?? new AppTheme();
            e.Graphics.Clear(activeTheme.ControlBack);

            using var borderPen = new Pen(activeTheme.Accent);
            using var textBrush = new SolidBrush(activeTheme.Fore);
            using var trackBrush = new System.Drawing.Drawing2D.LinearGradientBrush(
                GetScrollbarTrackRectangle(),
                Color.FromArgb(0, 28, 8),
                Color.FromArgb(0, 8, 2),
                System.Drawing.Drawing2D.LinearGradientMode.Horizontal);
            using var thumbBrush = new System.Drawing.Drawing2D.LinearGradientBrush(
                GetThumbRectangle(),
                Color.FromArgb(120, 255, 135),
                activeTheme.Accent,
                System.Drawing.Drawing2D.LinearGradientMode.Vertical);

            Rectangle content = GetContentRectangle();
            int lineHeight = GetLineHeight();
            int visibleLines = GetVisibleLineCount();

            for (int i = 0; i < visibleLines; i++)
            {
                int lineIndex = firstLine + i;
                if (lineIndex >= wrappedLines.Count)
                    break;

                TextRenderer.DrawText(
                    e.Graphics,
                    wrappedLines[lineIndex],
                    Font,
                    new Rectangle(content.X, content.Y + (i * lineHeight), content.Width, lineHeight),
                    activeTheme.Fore,
                    TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPrefix);
            }

            Rectangle track = GetScrollbarTrackRectangle();
            e.Graphics.FillRectangle(trackBrush, track);
            e.Graphics.FillRectangle(thumbBrush, GetThumbRectangle());
            e.Graphics.DrawRectangle(borderPen, 0, 0, Width - 1, Height - 1);
        }

        private void RebuildLines()
        {
            wrappedLines.Clear();
            int maxWidth = Math.Max(GetContentRectangle().Width - 6, 40);

            foreach (string sourceLine in text.Replace("\r\n", "\n").Split('\n'))
            {
                if (string.IsNullOrWhiteSpace(sourceLine))
                {
                    wrappedLines.Add("");
                    continue;
                }

                var current = "";
                foreach (string word in sourceLine.Split(' '))
                {
                    string candidate = string.IsNullOrEmpty(current) ? word : current + " " + word;
                    if (TextRenderer.MeasureText(candidate, Font).Width <= maxWidth)
                    {
                        current = candidate;
                    }
                    else
                    {
                        if (!string.IsNullOrEmpty(current))
                        {
                            wrappedLines.Add(current);
                        }

                        current = word;
                    }
                }

                wrappedLines.Add(current);
            }

            firstLine = Math.Min(firstLine, GetMaxFirstLine());
        }

        private void ScrollBy(int lineCount)
        {
            firstLine = Math.Max(0, Math.Min(GetMaxFirstLine(), firstLine + lineCount));
            Invalidate();
        }

        private Rectangle GetContentRectangle()
        {
            return new Rectangle(10, 8, Math.Max(Width - ScrollbarWidth - 24, 10), Math.Max(Height - 16, 10));
        }

        private Rectangle GetScrollbarTrackRectangle()
        {
            return new Rectangle(Width - ScrollbarWidth - 6, 6, ScrollbarWidth, Math.Max(Height - 12, 10));
        }

        private Rectangle GetThumbRectangle()
        {
            Rectangle track = GetScrollbarTrackRectangle();
            int totalLines = Math.Max(wrappedLines.Count, 1);
            int visibleLines = GetVisibleLineCount();

            if (totalLines <= visibleLines)
            {
                return new Rectangle(track.X + 2, track.Y + 2, track.Width - 4, track.Height - 4);
            }

            int thumbHeight = Math.Max(32, track.Height * visibleLines / totalLines);
            int maxFirstLine = GetMaxFirstLine();
            int movablePixels = Math.Max(track.Height - thumbHeight - 4, 1);
            int thumbY = track.Y + 2 + (maxFirstLine == 0 ? 0 : firstLine * movablePixels / maxFirstLine);

            return new Rectangle(track.X + 2, thumbY, track.Width - 4, thumbHeight);
        }

        private int GetVisibleLineCount()
        {
            return Math.Max(1, GetContentRectangle().Height / GetLineHeight());
        }

        private int GetLineHeight()
        {
            return Font.Height + 4;
        }

        private int GetMaxFirstLine()
        {
            return Math.Max(0, wrappedLines.Count - GetVisibleLineCount());
        }
    }
}
