using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Windows.Forms;
using ExcelTrainingMonitor.Controls;

namespace ExcelTrainingMonitor.Services
{
    internal sealed class AppTheme
    {
        public Color WindowBack { get; init; } = Color.Black;
        public Color PanelBack { get; init; } = Color.FromArgb(8, 10, 9);
        public Color ControlBack { get; init; } = Color.FromArgb(3, 45, 13);
        public Color ControlHover { get; init; } = Color.FromArgb(5, 78, 20);
        public Color Fore { get; init; } = Color.FromArgb(0, 255, 40);
        public Color MutedFore { get; init; } = Color.FromArgb(120, 255, 135);
        public Color Accent { get; init; } = Color.FromArgb(0, 220, 35);
        public Image BackgroundImage { get; init; }
        public Image LogoImage { get; init; }
        public Image AccentBarImage { get; init; }
    }

    internal static class ThemeManager
    {
        public static AppTheme LoadTheme(string themeName)
        {
            string selectedTheme = string.Equals(themeName, "Graphite", StringComparison.OrdinalIgnoreCase)
                ? "Graphite"
                : "Dark";
            string assetPath = Path.Combine(AppContext.BaseDirectory, "Themes", selectedTheme);
            if (!Directory.Exists(assetPath))
                assetPath = Path.Combine(AppContext.BaseDirectory, "Themes", "Dark");

            if (selectedTheme == "Graphite")
            {
                return new AppTheme
                {
                    WindowBack = Color.FromArgb(5, 5, 6),
                    PanelBack = Color.FromArgb(16, 17, 19),
                    ControlBack = Color.FromArgb(35, 37, 40),
                    ControlHover = Color.FromArgb(52, 56, 59),
                    Fore = Color.FromArgb(222, 225, 226),
                    MutedFore = Color.FromArgb(148, 154, 157),
                    Accent = Color.FromArgb(36, 205, 82),
                    BackgroundImage = LoadImage(Path.Combine(assetPath, "background.png")),
                    LogoImage = LoadImage(Path.Combine(assetPath, "logo.png")),
                    AccentBarImage = LoadImage(Path.Combine(assetPath, "divider.png"))
                };
            }

            return new AppTheme
            {
                BackgroundImage = LoadImage(Path.Combine(assetPath, "background.png")),
                LogoImage = LoadImage(Path.Combine(assetPath, "logo.png")),
                AccentBarImage = LoadImage(Path.Combine(assetPath, "divider.png"))
            };
        }

        private static Image LoadImage(string path)
        {
            if (!File.Exists(path))
                return null;

            using var stream = File.OpenRead(path);
            using var source = Image.FromStream(stream);
            return new Bitmap(source);
        }

        public static void Apply(Form form, AppTheme theme)
        {
            form.BackColor = theme.WindowBack;
            form.ForeColor = theme.Fore;
            form.BackgroundImage = null;

            ApplyControl(form, theme);
        }

        private static void ApplyControl(Control control, AppTheme theme)
        {
            if (control is Form)
            {
                control.BackColor = theme.WindowBack;
            }
            else if (control is TabPage || control is TableLayoutPanel || control is FlowLayoutPanel || control is Panel)
            {
                control.BackColor = theme.PanelBack;
                control.ForeColor = theme.Fore;
                control.Tag = theme;
                control.Paint -= DrawGlossyContainer;
                control.Paint += DrawGlossyContainer;

                if (control is TableLayoutPanel table)
                {
                    table.CellBorderStyle = TableLayoutPanelCellBorderStyle.None;
                }
            }
            else if (control is GlossyButton glossyButton)
            {
                glossyButton.BackColor = theme.ControlBack;
                glossyButton.ForeColor = theme.Fore;
                glossyButton.AccentColor = theme.Accent;
                glossyButton.HoverBackColor = theme.ControlHover;
                glossyButton.Font = new Font(glossyButton.Font, FontStyle.Regular);
            }
            else if (control is Button button)
            {
                button.FlatStyle = FlatStyle.Flat;
                button.FlatAppearance.BorderColor = theme.Accent;
                button.FlatAppearance.BorderSize = 1;
                button.FlatAppearance.MouseDownBackColor = ControlPaint.Dark(theme.ControlHover, 0.25F);
                button.FlatAppearance.MouseOverBackColor = theme.ControlHover;
                button.BackColor = theme.ControlBack;
                button.ForeColor = theme.Fore;
                button.BackgroundImage = null;
                button.Font = new Font(button.Font, FontStyle.Regular);
                button.UseVisualStyleBackColor = false;
            }
            else if (control is GlossyComboBox glossyComboBox)
            {
                glossyComboBox.BackColor = theme.ControlBack;
                glossyComboBox.ForeColor = theme.Fore;
                glossyComboBox.AccentColor = theme.Accent;
                glossyComboBox.HoverBackColor = theme.ControlHover;
                glossyComboBox.Tag = theme;
                glossyComboBox.FlatStyle = FlatStyle.Flat;
                glossyComboBox.DrawMode = DrawMode.OwnerDrawFixed;
                glossyComboBox.Invalidate();
            }
            else if (control is TextBox || control is NumericUpDown || control is DateTimePicker || control is ComboBox)
            {
                control.BackColor = theme.ControlBack;
                control.ForeColor = theme.Fore;

                if (control is TextBox textBox)
                {
                    textBox.BorderStyle = BorderStyle.FixedSingle;
                }
                else if (control is ComboBox comboBox)
                {
                    comboBox.FlatStyle = FlatStyle.Flat;
                    comboBox.DrawMode = DrawMode.OwnerDrawFixed;
                    comboBox.Tag = theme;
                    comboBox.DrawItem -= DrawThemedComboItem;
                    comboBox.DrawItem += DrawThemedComboItem;
                }
            }
            else if (control is TabControl tabControl)
            {
                tabControl.DrawMode = TabDrawMode.OwnerDrawFixed;
                tabControl.BackColor = theme.PanelBack;
                tabControl.Tag = theme;
                tabControl.DrawItem -= DrawThemedTab;
                tabControl.DrawItem += DrawThemedTab;
            }
            else if (control is GlossyCheckBox glossyCheckBox)
            {
                glossyCheckBox.BackColor = Color.Transparent;
                glossyCheckBox.ForeColor = theme.Fore;
                glossyCheckBox.AccentColor = theme.Accent;
                glossyCheckBox.BoxBackColor = theme.ControlBack;
                glossyCheckBox.HoverBackColor = theme.ControlHover;
                glossyCheckBox.Invalidate();
            }
            else if (control is CheckBox || control is Label)
            {
                control.BackColor = Color.Transparent;
                control.ForeColor = theme.Fore;

                if (control is CheckBox checkBox)
                {
                    checkBox.FlatStyle = FlatStyle.Flat;
                }
            }
            else if (control is DataGridView grid)
            {
                grid.BackgroundColor = theme.PanelBack;
                grid.Tag = theme;
                grid.EnableHeadersVisualStyles = false;
                grid.ColumnHeadersDefaultCellStyle.BackColor = theme.ControlBack;
                grid.ColumnHeadersDefaultCellStyle.ForeColor = theme.Fore;
                grid.DefaultCellStyle.BackColor = theme.PanelBack;
                grid.DefaultCellStyle.ForeColor = theme.Fore;
                grid.DefaultCellStyle.SelectionBackColor = theme.Accent;
                grid.DefaultCellStyle.SelectionForeColor = Color.Black;
                grid.BorderStyle = BorderStyle.FixedSingle;
                // disable built-in cell borders and draw semi-transparent lines in Paint
                grid.CellBorderStyle = DataGridViewCellBorderStyle.None;
                grid.Paint -= DrawTransparentGridLines;
                grid.Paint += DrawTransparentGridLines;
                ApplyTrainingColors(grid);
                grid.CellEndEdit += (_, __) => ApplyTrainingColors(grid);
            }
            else if (control is PictureBox pictureBox)
            {
                pictureBox.BackColor = Color.Transparent;
            }
            else
            {
                control.ForeColor = theme.Fore;
            }

            foreach (Control child in control.Controls)
            {
                ApplyControl(child, theme);
            }

        }

        private static void DrawThemedComboItem(object sender, DrawItemEventArgs e)
        {
            if (sender is not ComboBox comboBox || e.Index < 0)
                return;

            bool selected = (e.State & DrawItemState.Selected) == DrawItemState.Selected;
            Rectangle bounds = e.Bounds;
            var theme = comboBox.Tag as AppTheme ?? new AppTheme();

            using var backBrush = new LinearGradientBrush(
                bounds,
                selected ? theme.ControlHover : theme.ControlBack,
                selected ? ControlPaint.Dark(theme.ControlHover, 0.3F) : ControlPaint.Dark(theme.ControlBack, 0.25F),
                LinearGradientMode.Vertical);
            using var shineBrush = new LinearGradientBrush(
                new Rectangle(bounds.X, bounds.Y, bounds.Width, Math.Max(4, bounds.Height / 2)),
                Color.FromArgb(selected ? 80 : 35, theme.Accent),
                Color.FromArgb(0, theme.Accent),
                LinearGradientMode.Vertical);

            e.Graphics.FillRectangle(backBrush, bounds);
            e.Graphics.FillRectangle(shineBrush, bounds.X, bounds.Y, bounds.Width, Math.Max(4, bounds.Height / 2));

            TextRenderer.DrawText(
                e.Graphics,
                comboBox.GetItemText(comboBox.Items[e.Index]),
                comboBox.Font,
                new Rectangle(bounds.X + 8, bounds.Y, bounds.Width - 12, bounds.Height),
                selected ? theme.Fore : comboBox.ForeColor,
                TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis);
        }

        private static void DrawThemedTab(object sender, DrawItemEventArgs e)
        {
            var tabControl = (TabControl)sender;
            var theme = tabControl.Tag as AppTheme ?? new AppTheme();
            TabPage page = tabControl.TabPages[e.Index];
            bool selected = e.Index == tabControl.SelectedIndex;
            Rectangle bounds = e.Bounds;

            using var backBrush = new LinearGradientBrush(
                bounds,
                selected ? theme.ControlHover : theme.ControlBack,
                selected ? ControlPaint.Dark(theme.ControlHover, 0.35F) : theme.PanelBack,
                LinearGradientMode.Vertical);
            using var shineBrush = new LinearGradientBrush(
                new Rectangle(bounds.X + 1, bounds.Y + 1, Math.Max(1, bounds.Width - 2), Math.Max(1, bounds.Height / 2)),
                Color.FromArgb(selected ? 75 : 25, theme.Accent),
                Color.FromArgb(0, theme.Accent),
                LinearGradientMode.Vertical);
            using var borderPen = new Pen(selected ? theme.Accent : Color.FromArgb(90, theme.Accent));
            using var textBrush = new SolidBrush(theme.Fore);

            e.Graphics.FillRectangle(backBrush, bounds);
            e.Graphics.FillRectangle(shineBrush, bounds.X + 1, bounds.Y + 1, Math.Max(1, bounds.Width - 2), Math.Max(1, bounds.Height / 2));
            e.Graphics.DrawRectangle(borderPen, bounds.X, bounds.Y, bounds.Width - 1, bounds.Height - 1);

            TextRenderer.DrawText(
                e.Graphics,
                page.Text,
                tabControl.Font,
                bounds,
                theme.Fore,
                TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPrefix);
        }

        private static void DrawGlossyContainer(object sender, PaintEventArgs e)
        {
            var control = (Control)sender;
            if (control.Width <= 0 || control.Height <= 0)
                return;
            var theme = control.Tag as AppTheme ?? new AppTheme();

            Rectangle bounds = new Rectangle(0, 0, control.Width, control.Height);
            Rectangle shine = new Rectangle(0, 0, control.Width, Math.Max(12, control.Height / 3));

            using var backBrush = new LinearGradientBrush(
                bounds,
                theme.PanelBack,
                ControlPaint.Dark(theme.PanelBack, 0.35F),
                LinearGradientMode.Vertical);
            using var shineBrush = new LinearGradientBrush(
                shine,
                Color.FromArgb(38, theme.Accent),
                Color.FromArgb(0, theme.Accent),
                LinearGradientMode.Vertical);
            using var borderPen = new Pen(Color.FromArgb(100, theme.Accent));

            e.Graphics.FillRectangle(backBrush, bounds);
            e.Graphics.FillRectangle(shineBrush, shine);

            if (control is TabPage || control is Panel)
            {
                e.Graphics.DrawRectangle(borderPen, 0, 0, control.Width - 1, control.Height - 1);
            }
        }

        private static void DrawTransparentGridLines(object sender, PaintEventArgs e)
        {
            if (sender is not DataGridView dgv)
                return;

            var theme = dgv.Tag as AppTheme ?? new AppTheme();
            using var pen = new Pen(Color.FromArgb(110, theme.Accent), 1);

            // draw horizontal separators for visible rows
            for (int i = 0; i < dgv.Rows.Count; i++)
            {
                var rect = dgv.GetRowDisplayRectangle(i, true);
                if (rect.Height <= 0)
                    continue;

                int y = rect.Bottom - 1;
                e.Graphics.DrawLine(pen, rect.Left, y, rect.Right, y);
            }
        }

        public static void ApplyTrainingColors(DataGridView dgvGridBook)
        {
            foreach (DataGridViewRow row in dgvGridBook.Rows)
            {
                foreach (DataGridViewCell cell in row.Cells)
                {
                    if (cell.Value == null)
                        continue;

                    string value = cell.Value.ToString()?.Trim();

                    switch (value)
                    {
                        case "0":
                            cell.Style.BackColor = Color.FromArgb(80, 15, 15);
                            cell.Style.ForeColor = Color.FromArgb(255, 180, 180);
                            break;

                        case "1":
                            cell.Style.BackColor = Color.FromArgb(90, 70, 0);
                            cell.Style.ForeColor = Color.FromArgb(255, 230, 100);
                            break;

                        case "2":
                            cell.Style.BackColor = Color.FromArgb(15, 80, 15);
                            cell.Style.ForeColor = Color.FromArgb(120, 255, 135);
                            break;

                        default:
                            cell.Style.BackColor = dgvGridBook.DefaultCellStyle.BackColor;
                            cell.Style.ForeColor = dgvGridBook.DefaultCellStyle.ForeColor;
                            break;
                    }
                }
            }

            dgvGridBook.Refresh();
        }
    }
}
