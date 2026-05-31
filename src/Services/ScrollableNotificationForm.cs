using ExcelTrainingMonitor.Controls;
using System.Drawing;
using System.Windows.Forms;

namespace ExcelTrainingMonitor.Services
{
    internal sealed class ScrollableNotificationForm : Form
    {
        public ScrollableNotificationForm(string title, string message)
        {
            Text = title;
            StartPosition = FormStartPosition.CenterScreen;
            Size = new Size(620, 460);
            MinimumSize = new Size(420, 280);
            ShowIcon = false;
            ShowInTaskbar = false;

            var layout = new TableLayoutPanel
            {
                ColumnCount = 1,
                Dock = DockStyle.Fill,
                Padding = new Padding(14),
                RowCount = 2
            };

            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            AppTheme theme = ThemeManager.LoadTheme(SettingsManager.Load().ThemeName);

            var textPanel = new ThemedScrollTextPanel
            {
                Dock = DockStyle.Fill,
                Font = new Font("Segoe UI", 10F)
            };

            textPanel.SetTheme(theme);
            textPanel.SetText(message);

            var closeButton = new Button
            {
                Anchor = AnchorStyles.Right,
                AutoSize = true,
                DialogResult = DialogResult.OK,
                Margin = new Padding(0, 12, 0, 0),
                Text = "OK"
            };

            layout.Controls.Add(textPanel, 0, 0);
            layout.Controls.Add(closeButton, 0, 1);
            Controls.Add(layout);
            AcceptButton = closeButton;

            ThemeManager.Apply(this, theme);
        }
    }
}
