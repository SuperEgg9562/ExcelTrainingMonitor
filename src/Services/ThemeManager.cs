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
            string assetPath = Path.Combine(AppContext.BaseDirectory, "Themes", themeName ?? "Dark");

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
                glossyButton.Font = new Font(glossyButton.Font, FontStyle.Regular);
            }
            else if (control is Button button)
            {
                button.FlatStyle = FlatStyle.Flat;
                button.FlatAppearance.BorderColor = theme.Accent;
                button.FlatAppearance.BorderSize = 1;
                button.FlatAppearance.MouseDownBackColor = Color.FromArgb(22, 100, 28);
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
                grid.GridColor = Color.FromArgb(0, 75, 18);
                grid.EnableHeadersVisualStyles = false;
                grid.ColumnHeadersDefaultCellStyle.BackColor = theme.ControlBack;
                grid.ColumnHeadersDefaultCellStyle.ForeColor = theme.Fore;
                grid.DefaultCellStyle.BackColor = theme.PanelBack;
                grid.DefaultCellStyle.ForeColor = theme.Fore;
                grid.DefaultCellStyle.SelectionBackColor = theme.Accent;
                grid.DefaultCellStyle.SelectionForeColor = Color.Black;
                grid.BorderStyle = BorderStyle.FixedSingle;
                grid.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal;
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

            using var backBrush = new LinearGradientBrush(
                bounds,
                selected ? Color.FromArgb(18, 96, 18) : Color.FromArgb(5, 48, 12),
                selected ? Color.FromArgb(0, 140, 28) : Color.FromArgb(0, 18, 6),
                LinearGradientMode.Vertical);
            using var shineBrush = new LinearGradientBrush(
                new Rectangle(bounds.X, bounds.Y, bounds.Width, Math.Max(4, bounds.Height / 2)),
                Color.FromArgb(selected ? 120 : 80, 255, 236, 125),
                Color.FromArgb(0, 0, 255, 40),
                LinearGradientMode.Vertical);

            e.Graphics.FillRectangle(backBrush, bounds);
            e.Graphics.FillRectangle(shineBrush, bounds.X, bounds.Y, bounds.Width, Math.Max(4, bounds.Height / 2));

            TextRenderer.DrawText(
                e.Graphics,
                comboBox.GetItemText(comboBox.Items[e.Index]),
                comboBox.Font,
                new Rectangle(bounds.X + 8, bounds.Y, bounds.Width - 12, bounds.Height),
                selected ? Color.FromArgb(255, 220, 20) : comboBox.ForeColor,
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
                selected ? Color.FromArgb(8, 92, 24) : Color.FromArgb(12, 18, 14),
                selected ? Color.FromArgb(0, 38, 10) : theme.PanelBack,
                LinearGradientMode.Vertical);
            using var shineBrush = new LinearGradientBrush(
                new Rectangle(bounds.X + 1, bounds.Y + 1, Math.Max(1, bounds.Width - 2), Math.Max(1, bounds.Height / 2)),
                selected ? Color.FromArgb(120, 180, 255, 185) : Color.FromArgb(45, 100, 150, 105),
                Color.FromArgb(0, 0, 255, 40),
                LinearGradientMode.Vertical);
            using var borderPen = new Pen(selected ? theme.Accent : Color.FromArgb(0, 80, 18));
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

            Rectangle bounds = new Rectangle(0, 0, control.Width, control.Height);
            Rectangle shine = new Rectangle(0, 0, control.Width, Math.Max(12, control.Height / 3));

            using var backBrush = new LinearGradientBrush(
                bounds,
                Color.FromArgb(4, 24, 8),
                Color.FromArgb(0, 6, 2),
                LinearGradientMode.Vertical);
            using var shineBrush = new LinearGradientBrush(
                shine,
                Color.FromArgb(60, 0, 255, 48),
                Color.FromArgb(0, 0, 80, 12),
                LinearGradientMode.Vertical);
            using var borderPen = new Pen(Color.FromArgb(0, 90, 18));

            e.Graphics.FillRectangle(backBrush, bounds);
            e.Graphics.FillRectangle(shineBrush, shine);

            if (control is TabPage || control is Panel)
            {
                e.Graphics.DrawRectangle(borderPen, 0, 0, control.Width - 1, control.Height - 1);
            }
        }
    }
}
