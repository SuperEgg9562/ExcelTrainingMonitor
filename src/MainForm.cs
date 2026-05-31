using ExcelTrainingMonitor.Models;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using ExcelTrainingMonitor.Services;
using System.Windows.Forms;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using ExcelTrainingMonitor.Controls;

namespace ExcelTrainingMonitor
{
    public partial class MainForm : Form
    {
        private string excelPath = "";
        private FileSystemWatcher watcher;
        private System.Windows.Forms.Timer scanTimer;
        private DateTime lastScan = DateTime.MinValue;
        private DateTime lastReminderShownFor = DateTime.MinValue;
        private AppSettings appSettings;
        private List<TrainingAlert> currentAlerts = new List<TrainingAlert>();
        private List<TrainingAlert> previousAlerts = new List<TrainingAlert>();

        [SupportedOSPlatform("windows")]
        [DllImport("user32.dll")]
        static extern bool SetForegroundWindow(IntPtr hWnd);

        [SupportedOSPlatform("windows")]
        [DllImport("user32.dll")]
        static extern bool ReleaseCapture();

        [SupportedOSPlatform("windows")]
        [DllImport("user32.dll")]
        static extern IntPtr SendMessage(IntPtr hWnd, int msg, int wParam, int lParam);

        public MainForm()
        {
            InitializeComponent();

            appSettings = SettingsManager.Load();
            excelPath = appSettings.ExcelPath;
            LoadSettingsIntoControls();
            WireSettingsEvents();
            ApplyTheme();

            if (!string.IsNullOrWhiteSpace(excelPath))
            {
                lblFile.Text = excelPath;
            }

            if (File.Exists(excelPath))
            {
                StartWatcher();
            }

            this.Resize += MainForm_Resize;

            dgvAlerts.CellFormatting += DgvAlerts_CellFormatting;
            ConfigureGrid(dgvAlerts);

            ConfigureGrid(dgvHistory);
            RefreshHistoryGrid();
            UpdateScanTimer();
        }

        private void LoadSettingsIntoControls()
        {
            numScanHours.Value = Math.Min(numScanHours.Maximum, Math.Max(numScanHours.Minimum, appSettings.ScanIntervalHours));
            numScanMinutes.Value = Math.Min(numScanMinutes.Maximum, Math.Max(numScanMinutes.Minimum, appSettings.ScanIntervalMinutes));
            chkReminderEnabled.Checked = appSettings.ReminderEnabled;
            dtpReminderDate.Value = appSettings.ReminderDateTime < dtpReminderDate.MinDate
                ? DateTime.Now.AddDays(1)
                : appSettings.ReminderDateTime;

            if (!cboTheme.Items.Contains(appSettings.ThemeName))
            {
                cboTheme.Items.Add(appSettings.ThemeName);
            }

            cboTheme.SelectedItem = appSettings.ThemeName;
            if (cboTheme.SelectedIndex < 0)
            {
                cboTheme.SelectedIndex = 0;
            }
        }

        private void WireSettingsEvents()
        {
            numScanHours.ValueChanged += SettingsControl_Changed;
            numScanMinutes.ValueChanged += SettingsControl_Changed;
            chkReminderEnabled.CheckedChanged += SettingsControl_Changed;
            dtpReminderDate.ValueChanged += SettingsControl_Changed;
            cboTheme.SelectedIndexChanged += SettingsControl_Changed;
        }

        private void SettingsControl_Changed(object sender, EventArgs e)
        {
            SaveSettingsFromControls();
            UpdateScanTimer();

            if (sender == cboTheme)
            {
                ApplyTheme();
            }
        }

        private void SaveSettingsFromControls()
        {
            appSettings.ExcelPath = excelPath;
            appSettings.ScanIntervalHours = (int)numScanHours.Value;
            appSettings.ScanIntervalMinutes = (int)numScanMinutes.Value;
            appSettings.ReminderEnabled = chkReminderEnabled.Checked;
            appSettings.ReminderDateTime = dtpReminderDate.Value;
            appSettings.ThemeName = cboTheme.SelectedItem?.ToString() ?? "Dark";
            SettingsManager.Save(appSettings);
        }

        private void ApplyTheme()
        {
            AppTheme theme = ThemeManager.LoadTheme(cboTheme.SelectedItem?.ToString() ?? "Dark");
            ThemeManager.Apply(this, theme);
            picThemeLogo.Image = theme.LogoImage;
            picAccentBar.Image = theme.AccentBarImage;
            picAccentBar.BackColor = theme.Accent;
            lblWindowTitle.ForeColor = theme.Fore;
            ConfigureGrid(dgvHistory);
            dgvAlerts.Refresh();
            dgvHistory.Refresh();
            statusPieChart.Refresh();
            openPieChart.Refresh();
            Invalidate();
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);

            using var borderPen = new Pen(Color.FromArgb(0, 220, 35), 2F);
            e.Graphics.DrawRectangle(borderPen, 1, 1, Width - 3, Height - 3);
        }

        private void TitleBar_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Left || !OperatingSystem.IsWindows())
                return;

            ReleaseCapture();
            SendMessage(Handle, 0xA1, 0x2, 0);
        }

        private void btnWindowMinimize_Click(object sender, EventArgs e)
        {
            WindowState = FormWindowState.Minimized;
        }

        private void btnWindowMaximize_Click(object sender, EventArgs e)
        {
            WindowState = WindowState == FormWindowState.Maximized
                ? FormWindowState.Normal
                : FormWindowState.Maximized;
        }

        private void btnWindowClose_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();

            dialog.Filter = "Excel Files (*.xlsx)|*.xlsx";

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                excelPath = dialog.FileName;

                lblFile.Text = excelPath;

                SettingsManager.SaveExcelPath(excelPath);
                appSettings.ExcelPath = excelPath;
                SaveSettingsFromControls();

                if (File.Exists(excelPath))
                {
                    StartWatcher();
                }
            }
        }

        private void btnStart_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(excelPath))
            {
                MessageBox.Show("Select an Excel file first.");

                return;
            }

            StartWatcher();

            NotificationManager.ShowNotification("Excel Monitor", "Live Monitoring Started");

            RunScan();
        }

        private void btnStop_Click(object sender, EventArgs e)
        {
            if (watcher != null)
            {
                watcher.EnableRaisingEvents = false;

                watcher.Dispose();

                watcher = null;
            }

            scanTimer?.Stop();

            NotificationManager.ShowNotification("Excel Monitor", "Monitoring Stopped");
        }
        private void RunScan(bool showUpdateNotification = false)
        {
            if (string.IsNullOrWhiteSpace(excelPath))
                return;

            List<TrainingAlert> alerts = ExcelMonitor.ScanFile(excelPath);
            List<TrainingChange> changes = GetTrainingChanges(alerts);

            UpdateDashboard(alerts);

            foreach (var current in alerts)
            {
                var previous =
                    previousAlerts.FirstOrDefault(x =>
                        x.EmployeeName == current.EmployeeName &&
                        x.Category == current.Category);

                if (previous != null)
                {
                    if (previous.Status != current.Status)
                    {
                        HistoryManager.Add(
                            current.EmployeeName,
                            current.Category,
                            previous.Status,
                            current.Status);
                    }
                }
            }

            previousAlerts =
                alerts
                    .Select(x => new TrainingAlert
                    {
                        EmployeeName = x.EmployeeName,
                        Category = x.Category,
                        Status = x.Status,
                        Timestamp = x.Timestamp
                    })
                    .ToList();

            dgvAlerts.DataSource = null;
            dgvAlerts.DataSource = alerts;

            currentAlerts = alerts;
            UpdateCharts(alerts);

            List<TrainingAlert> newAlerts = AlertStateManager.GetNewAlerts(alerts);
            List<TrainingAlert> openNewAlerts =
                newAlerts
                    .Where(x => IsOpenTrainingStatus(x.Status))
                    .ToList();

            if (showUpdateNotification)
            {
                ShowExcelUpdateNotification(changes);
            }
            else if (openNewAlerts.Count > 0)
            {
                ShowOpenTrainingNotification(openNewAlerts);
                ForceForeground();
            }

            if (ReminderIsDue())
            {
                ShowDateReminderNotification();
            }

            lblTotal.Text = "Total Flags: " + alerts.Count;
            lblNotTrained.Text = "Not Trained: " + alerts.Count(x => x.Status == "Not Trained");
            lblTraining.Text = "In Training: " + alerts.Count(x => x.Status == "In Training");
            lblComplete.Text = "Complete: " + alerts.Count(x => x.Status == "Complete");

            RefreshHistoryGrid();

        }

        private List<TrainingChange> GetTrainingChanges(List<TrainingAlert> alerts)
        {
            var changes = new List<TrainingChange>();

            foreach (var current in alerts)
            {
                var previous =
                    previousAlerts.FirstOrDefault(x =>
                        x.EmployeeName == current.EmployeeName &&
                        x.Category == current.Category);

                if (previous == null)
                {
                    changes.Add(new TrainingChange
                    {
                        EmployeeName = current.EmployeeName,
                        Category = current.Category,
                        OldStatus = "New",
                        NewStatus = current.Status,
                        Timestamp = DateTime.Now
                    });
                }
                else if (previous.Status != current.Status)
                {
                    changes.Add(new TrainingChange
                    {
                        EmployeeName = current.EmployeeName,
                        Category = current.Category,
                        OldStatus = previous.Status,
                        NewStatus = current.Status,
                        Timestamp = DateTime.Now
                    });
                }
            }

            return changes;
        }

        private void ShowOpenTrainingNotification(List<TrainingAlert> alerts)
        {
            var message = new StringBuilder();
            message.AppendLine($"{alerts.Count} open training item(s) need attention.");
            message.AppendLine();

            foreach (var alert in alerts.OrderBy(x => x.EmployeeName).ThenBy(x => x.Category))
            {
                message.AppendLine($"{alert.EmployeeName}");
                message.AppendLine($"  {alert.Category}");
                message.AppendLine($"  Status: {alert.Status}");
                message.AppendLine();
            }

            NotificationManager.ShowNotification("Training Alerts", message.ToString().TrimEnd());
        }

        private void ShowExcelUpdateNotification(List<TrainingChange> changes)
        {
            if (changes.Count == 0)
            {
                NotificationManager.ShowNotification(
                    "Excel Updated",
                    $"Updated: {DateTime.Now:yyyy-MM-dd HH:mm:ss}\r\nNo training status changes were detected.");
                return;
            }

            var message = new StringBuilder();
            message.AppendLine($"Updated: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            message.AppendLine($"{changes.Count} training item(s) changed.");
            message.AppendLine();

            foreach (var change in changes.OrderBy(x => x.EmployeeName).ThenBy(x => x.Category))
            {
                message.AppendLine($"{change.EmployeeName}");
                message.AppendLine($"  {change.Category}");
                message.AppendLine($"  {change.OldStatus} -> {change.NewStatus}");
                message.AppendLine();
            }

            NotificationManager.ShowNotification("Excel Updated", message.ToString().TrimEnd());
        }

        private static bool IsOpenTrainingStatus(string status)
        {
            return status == "Not Trained" || status == "In Training";
        }

        private bool ReminderIsDue()
        {
            if (!chkReminderEnabled.Checked)
                return false;

            DateTime reminderDate = dtpReminderDate.Value;

            if (DateTime.Now < reminderDate)
                return false;

            if (lastReminderShownFor == reminderDate)
                return false;

            return true;
        }

        private void ShowDateReminderNotification()
        {
            lastReminderShownFor = dtpReminderDate.Value;
            NotificationManager.ShowNotification(
                "Reminder",
                $"Reminder due: {dtpReminderDate.Value:yyyy-MM-dd HH:mm}\r\nExcel file: {excelPath}");
        }

        private void ConfigureGrid(DataGridView grid)
        {
            grid.EnableHeadersVisualStyles = false;
            grid.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(3, 45, 13);
            grid.ColumnHeadersDefaultCellStyle.ForeColor = Color.FromArgb(0, 255, 40);
            grid.BackgroundColor = Color.FromArgb(8, 10, 9);
            grid.DefaultCellStyle.BackColor = Color.FromArgb(8, 10, 9);
            grid.DefaultCellStyle.ForeColor = Color.FromArgb(0, 255, 40);
            grid.DefaultCellStyle.SelectionBackColor = Color.FromArgb(0, 220, 35);
            grid.DefaultCellStyle.SelectionForeColor = Color.Black;
            grid.GridColor = Color.FromArgb(0, 75, 18);
            grid.BorderStyle = BorderStyle.FixedSingle;
            grid.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal;
            grid.RowHeadersVisible = false;
            grid.ReadOnly = true;
            grid.AllowUserToAddRows = false;
            grid.AllowUserToDeleteRows = false;
            grid.AllowUserToResizeRows = false;
            grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        }

        private void RefreshHistoryGrid()
        {
            dgvHistory.DataSource = null;
            dgvHistory.DataSource = HistoryManager.GetHistory();
        }

        private void StartWatcher()
        {
            if (watcher != null)
            {
                watcher.EnableRaisingEvents = false;
                watcher.Dispose();
            }

            watcher = new FileSystemWatcher();
            watcher.Path = Path.GetDirectoryName(excelPath);
            watcher.Filter = Path.GetFileName(excelPath);
            watcher.NotifyFilter = NotifyFilters.LastWrite | NotifyFilters.Size | NotifyFilters.FileName;
            watcher.Changed += Watcher_Changed;
            watcher.Created += Watcher_Changed;
            watcher.Renamed += Watcher_Changed;
            watcher.EnableRaisingEvents = true;
            UpdateScanTimer();
        }

        private void UpdateScanTimer()
        {
            int intervalMinutes =
                ((int)numScanHours.Value * 60) +
                (int)numScanMinutes.Value;

            if (intervalMinutes < 1)
            {
                intervalMinutes = 1;
                numScanMinutes.Value = 1;
            }

            scanTimer ??= new System.Windows.Forms.Timer();
            scanTimer.Stop();
            scanTimer.Interval = intervalMinutes * 60 * 1000;
            scanTimer.Tick -= ScanTimer_Tick;
            scanTimer.Tick += ScanTimer_Tick;

            if (!string.IsNullOrWhiteSpace(excelPath) && File.Exists(excelPath) && watcher != null)
            {
                scanTimer.Start();
            }
        }

        private void ScanTimer_Tick(object sender, EventArgs e)
        {
            RunScan();
        }

        private void Watcher_Changed(
        object sender,
        FileSystemEventArgs e)
        {
            try
            {
                if ((DateTime.Now - lastScan).TotalSeconds < 3)
                    return;

                lastScan = DateTime.Now;
                System.Threading.Thread.Sleep(1000);
                this.Invoke(new Action(() =>
                {
                    RunScan(showUpdateNotification: true);
                }));
            }
            catch (Exception ex)
            {
                System.Diagnostics.Trace.WriteLine($"Watcher_Changed error: {ex}");
            }
        }
        private void DgvAlerts_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.RowIndex < 0)
                return;

            DataGridViewRow row = dgvAlerts.Rows[e.RowIndex];

            if (row.Cells["Status"].Value == null)
                return;

            string status = row.Cells["Status"].Value.ToString();

            row.DefaultCellStyle.SelectionForeColor = Color.Black;

            if (status == "Not Trained")
            {
                row.DefaultCellStyle.BackColor = Color.FromArgb(38, 0, 4);
                row.DefaultCellStyle.ForeColor = Color.FromArgb(255, 70, 70);
                row.DefaultCellStyle.SelectionBackColor = Color.FromArgb(180, 18, 22);
            }
            else if (status == "In Training")
            {
                row.DefaultCellStyle.BackColor = Color.FromArgb(44, 30, 0);
                row.DefaultCellStyle.ForeColor = Color.FromArgb(255, 210, 28);
                row.DefaultCellStyle.SelectionBackColor = Color.FromArgb(190, 130, 8);
            }
            else if (status == "Complete")
            {
                row.DefaultCellStyle.BackColor = Color.FromArgb(0, 20, 5);
                row.DefaultCellStyle.ForeColor = Color.FromArgb(0, 150, 25);
                row.DefaultCellStyle.SelectionBackColor = Color.FromArgb(0, 120, 22);
            }
        }

        private void ForceForeground()
        {
            this.WindowState = FormWindowState.Normal;
            this.Show();
            this.Activate();
            if (OperatingSystem.IsWindows())
            {
                SetForegroundWindow(this.Handle);
            }
        }

        private void MainForm_Resize(object sender, EventArgs e)
        {
            if (chkMinimizeTray.Checked &&
                this.WindowState == FormWindowState.Minimized)
            {
                this.Hide();

                notifyIcon1.BalloonTipTitle = "Excel Training Monitor";
                notifyIcon1.BalloonTipText = "Still running in background.";
                notifyIcon1.ShowBalloonTip(2000);
            }
        }

        private void notifyIcon1_DoubleClick(
            object sender,
            EventArgs e)
        {
            this.Show();
            this.WindowState = FormWindowState.Normal;
            ForceForeground();
        }

        private void TxtSearch_TextChanged(object sender,EventArgs e)
        {
            string search =
                txtSearch.Text.ToLower();

            dgvAlerts.DataSource = currentAlerts
                .Where(x =>
                    (x.EmployeeName ?? "").ToLower().Contains(search) ||
                    (x.Category ?? "").ToLower().Contains(search))
                .ToList();
        }
        private void UpdateDashboard(List<TrainingAlert> alerts)
        {
            lblTotal.Text = $"Total: {alerts.Count}";
            lblNotTrained.Text = $"Not Trained: {alerts.Count(a => a.Status == "Not Trained")}";
            lblTraining.Text = $"In Training: {alerts.Count(a => a.Status == "In Training")}";
            lblComplete.Text = $"Complete: {alerts.Count(a => a.Status == "Complete")}";

            pbNotTrained.Maximum = Math.Max(alerts.Count, 1);
            pbTraining.Maximum = Math.Max(alerts.Count, 1);
            pbComplete.Maximum = Math.Max(alerts.Count, 1);
            pbNotTrained.Value = alerts.Count(a => a.Status == "Not Trained");
            pbTraining.Value = alerts.Count(a => a.Status == "In Training");
            pbComplete.Value = alerts.Count(a => a.Status == "Complete");
        }

        private void UpdateCharts(List<TrainingAlert> alerts)
        {
            int notTrained = alerts.Count(a => a.Status == "Not Trained");
            int inTraining = alerts.Count(a => a.Status == "In Training");
            int complete = alerts.Count(a => a.Status == "Complete");

            statusPieChart.SetSegments(new[]
            {
                new PieSegment { Label = "Not Trained", Value = notTrained, Color = Color.FromArgb(220, 24, 30) },
                new PieSegment { Label = "In Training", Value = inTraining, Color = Color.FromArgb(255, 196, 22) },
                new PieSegment { Label = "Complete", Value = complete, Color = Color.FromArgb(0, 220, 35) }
            });

            openPieChart.SetSegments(new[]
            {
                new PieSegment { Label = "Not Trained", Value = notTrained, Color = Color.FromArgb(220, 24, 30) },
                new PieSegment { Label = "In Training", Value = inTraining, Color = Color.FromArgb(255, 196, 22) }
            });
        }

        private sealed class TrainingChange
        {
            public string EmployeeName { get; set; }
            public string Category { get; set; }
            public string OldStatus { get; set; }
            public string NewStatus { get; set; }
            public DateTime Timestamp { get; set; }
        }
    }
}
