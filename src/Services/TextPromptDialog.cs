using System.Drawing;
using System.Windows.Forms;

namespace ExcelTrainingMonitor.Services
{
    internal static class TextPromptDialog
    {
        public static bool Show(
            IWin32Window owner,
            string title,
            string prompt,
            string initialValue,
            bool multiline,
            out string value)
        {
            using var form = new Form
            {
                AcceptButton = null,
                CancelButton = null,
                ClientSize = new Size(520, multiline ? 390 : 160),
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false,
                MinimizeBox = false,
                ShowInTaskbar = false,
                StartPosition = FormStartPosition.CenterParent,
                Text = title
            };

            var label = new Label
            {
                AutoSize = true,
                Location = new Point(14, 14),
                Text = prompt
            };
            var textBox = new TextBox
            {
                AcceptsReturn = multiline,
                Location = new Point(14, 40),
                Multiline = multiline,
                ScrollBars = multiline ? ScrollBars.Vertical : ScrollBars.None,
                Size = new Size(492, multiline ? 292 : 28),
                Text = initialValue ?? ""
            };
            var ok = new Button
            {
                DialogResult = DialogResult.OK,
                Location = new Point(350, multiline ? 344 : 110),
                Size = new Size(75, 30),
                Text = "OK"
            };
            var cancel = new Button
            {
                DialogResult = DialogResult.Cancel,
                Location = new Point(431, multiline ? 344 : 110),
                Size = new Size(75, 30),
                Text = "Cancel"
            };

            form.Controls.AddRange(new Control[] { label, textBox, ok, cancel });
            form.AcceptButton = ok;
            form.CancelButton = cancel;
            textBox.SelectAll();

            bool accepted = form.ShowDialog(owner) == DialogResult.OK;
            value = textBox.Text.Trim();
            return accepted;
        }
    }
}
