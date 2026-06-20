using ExcelTrainingMonitor.Models;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
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
        private BindingList<TrainingAlert> currentAlertBinding = new BindingList<TrainingAlert>();
        private DataGridView dgvGridBook;
        private DataGridView dgvCompliancePlan;
        private DataGridView dgvProcessRecord;
        private DataGridView dgvDailyProduction;
        private TextBox txtComplianceTerms;
        private TextBox txtComplianceTitle;
        private TextBox txtComplianceLegend;
        private TextBox txtComplianceDetailIssues;
        private DateTimePicker dtpComplianceDateTime;
        private DateTimePicker dtpComplianceTime;
        private PictureBox picComplianceLogo;
        private Label lblComplianceLogoPlaceholder;
        private TextBox txtProcessRecordVersion;
        private TextBox txtProcessRecordTitle;
        private DateTimePicker dtpProcessRecordDateTime;
        private DateTimePicker dtpProcessRecordTime;
        private PictureBox picProcessRecordLogo;
        private Label lblProcessRecordLogoPlaceholder;
        private GlossyComboBox cboProcessSupplier;
        private NumericUpDown numProcessBirds;
        private GlossyComboBox cboProcessDropdownLists;
        private GlossyComboBox cboGridBookSheets;
        private GlossyComboBox cboMinimizeBehavior;
        private GlossyCheckBox chkReminderAgentOnClose;
        private DataTable currentGridBookTable = new DataTable();
        private DataTable compliancePlanTable = new DataTable();
        private DataTable processRecordTable = new DataTable();
        private DataTable dailyProductionTable = new DataTable();
        private ProcessRecordMetadata processRecordMetadata = new ProcessRecordMetadata();
        private DataGridView activeProcessGrid;
        private string currentGridBookSheet = "Sheet1";
        private string compliancePlanPath = "";
        private string processRecordPath = "";

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
            CreateGridBookEditorTab();
            CreateCompliancePlanTab();
            CreateProcessRecordTab();
            CreateBehaviorControls();
            CreateChartExportControls();
            ApplyFocusedTabLayout();

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
                LoadGridBookEditor(excelPath);
            }

            this.Resize += MainForm_Resize;

            dgvAlerts.CellFormatting += DgvAlerts_CellFormatting;
            ConfigureGrid(dgvAlerts);
            dgvAlerts.ReadOnly = false;
            dgvAlerts.AllowUserToAddRows = true;
            dgvAlerts.AllowUserToDeleteRows = true;

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

            chkNtfyEnabled.Checked = appSettings.NtfyEnabled;
            txtNtfyTopic.Text = appSettings.NtfyTopic;
            txtNtfyEmail.Text = appSettings.NtfyEmail;
            cboMinimizeBehavior.SelectedItem = appSettings.MinimizeBehavior == "Window" ? "Window" : "Tray";
            chkReminderAgentOnClose.Checked = appSettings.StartReminderAgentOnClose;
            RefreshProcessSupplierChoices();
        }

        private void WireSettingsEvents()
        {
            numScanHours.ValueChanged += SettingsControl_Changed;
            numScanMinutes.ValueChanged += SettingsControl_Changed;
            chkReminderEnabled.CheckedChanged += SettingsControl_Changed;
            dtpReminderDate.ValueChanged += SettingsControl_Changed;
            cboTheme.SelectedIndexChanged += SettingsControl_Changed;
            chkNtfyEnabled.CheckedChanged += SettingsControl_Changed;
            txtNtfyTopic.TextChanged += SettingsControl_Changed;
            txtNtfyEmail.TextChanged += SettingsControl_Changed;
            cboMinimizeBehavior.SelectedIndexChanged += SettingsControl_Changed;
            chkReminderAgentOnClose.CheckedChanged += SettingsControl_Changed;
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
            appSettings.NtfyEnabled = chkNtfyEnabled.Checked;
            appSettings.NtfyTopic = txtNtfyTopic.Text.Trim();
            appSettings.NtfyEmail = txtNtfyEmail.Text.Trim();
            appSettings.MinimizeBehavior = cboMinimizeBehavior.SelectedItem?.ToString() ?? "Tray";
            appSettings.StartReminderAgentOnClose = chkReminderAgentOnClose.Checked;
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
            dgvGridBook?.Refresh();
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

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            SaveSettingsFromControls();

            if (appSettings.StartReminderAgentOnClose && appSettings.ReminderEnabled)
            {
                StartReminderAgent();
            }

            base.OnFormClosing(e);
        }

        private void StartReminderAgent()
        {
            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = Application.ExecutablePath,
                    Arguments = "--reminder-agent",
                    UseShellExecute = false,
                    CreateNoWindow = true
                });
            }
            catch (Exception ex)
            {
                System.Diagnostics.Trace.WriteLine($"Could not start reminder agent: {ex}");
            }
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
                    LoadGridBookEditor(excelPath);
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

            currentAlerts = alerts;
            currentAlertBinding = new BindingList<TrainingAlert>(alerts);
            dgvAlerts.DataSource = null;
            dgvAlerts.DataSource = currentAlertBinding;
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

            _ = SendNtfyReminderAsync();
        }

        private async System.Threading.Tasks.Task SendNtfyReminderAsync()
        {
            try
            {
                SaveSettingsFromControls();
                await NtfyService.SendReminderAsync(appSettings, GetEditableAlerts());
            }
            catch (Exception ex)
            {
                System.Diagnostics.Trace.WriteLine($"ntfy reminder failed: {ex}");
            }
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

        private void CreateGridBookEditorTab()
        {
            var tabEditor = new TabPage
            {
                Name = "tabGridBookEditor",
                Text = "GridBook Editor",
                Padding = new Padding(3)
            };

            var editorLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 2,
                Margin = new Padding(0)
            };
            editorLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            editorLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            var toolbar = new FlowLayoutPanel
            {
                AutoSize = true,
                Dock = DockStyle.Fill,
                Margin = new Padding(0, 0, 0, 8),
                WrapContents = true
            };

            GlossyButton btnEditorNew = CreateActionButton("New GridBook", btnEditorNew_Click, 124);
            GlossyButton btnEditorOpen = CreateActionButton("Open GridBook", btnEditorOpen_Click, 130);
            GlossyButton btnEditorSave = CreateActionButton("Save Sheet", btnEditorSave_Click, 108);
            GlossyButton btnEditorSaveAs = CreateActionButton("Save As", btnEditorSaveAs_Click, 96);
            GlossyButton btnEditorExport = CreateActionButton("Export Copy", btnEditorExport_Click, 112);
            GlossyButton btnEditorAddSheet = CreateActionButton("Add Sheet", btnEditorAddSheet_Click, 104);
            GlossyButton btnEditorAddRow = CreateActionButton("Add Row", btnEditorAddRow_Click, 92);
            GlossyButton btnEditorAddColumn = CreateActionButton("Add Column", btnEditorAddColumn_Click, 116);
            GlossyButton btnEditorTrainingColors = CreateActionButton("Training Colors", btnEditorTrainingColors_Click, 132);

            cboGridBookSheets = new GlossyComboBox
            {
                Name = "cboGridBookSheets",
                Width = 200,
                Margin = new Padding(0, 0, 8, 6)
            };
            cboGridBookSheets.SelectedIndexChanged += cboGridBookSheets_SelectedIndexChanged;

            toolbar.Controls.Add(btnEditorNew);
            toolbar.Controls.Add(btnEditorOpen);
            toolbar.Controls.Add(cboGridBookSheets);
            toolbar.Controls.Add(btnEditorAddSheet);
            toolbar.Controls.Add(btnEditorAddRow);
            toolbar.Controls.Add(btnEditorAddColumn);
            toolbar.Controls.Add(btnEditorSave);
            toolbar.Controls.Add(btnEditorSaveAs);
            toolbar.Controls.Add(btnEditorExport);
            toolbar.Controls.Add(btnEditorTrainingColors);

            dgvGridBook = new DataGridView
            {
                Dock = DockStyle.Fill,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None,
                AllowUserToAddRows = true,
                AllowUserToDeleteRows = true,
                RowHeadersWidth = 52,
                ColumnHeadersHeight = 32,
                Name = "dgvGridBook"
            };
            dgvGridBook.RowTemplate.Height = 28;
            dgvGridBook.RowPostPaint += DgvGridBook_RowPostPaint;
            ConfigureGrid(dgvGridBook);
            dgvGridBook.ReadOnly = false;
            dgvGridBook.AllowUserToAddRows = true;
            dgvGridBook.AllowUserToDeleteRows = true;

            editorLayout.Controls.Add(toolbar, 0, 0);
            editorLayout.Controls.Add(dgvGridBook, 0, 1);
            tabEditor.Controls.Add(editorLayout);
            tabControl1.Controls.Add(tabEditor);
            tabControl1.SelectedIndexChanged += TabControl1_SelectedIndexChanged;
        }

        private void CreateCompliancePlanTab()
        {
            var tabCompliance = new TabPage
            {
                Name = "tabCompliancePlan",
                Text = "Compliance Plan",
                Padding = new Padding(3)
            };

            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 7,
                Margin = new Padding(0)
            };
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            txtComplianceTerms = new TextBox
            {
                AcceptsReturn = true,
                AutoSize = false,
                Dock = DockStyle.Fill,
                Height = 46,
                Margin = new Padding(0, 0, 0, 6),
                Multiline = true,
                PlaceholderText = "Technical terms",
                WordWrap = true
            };
            txtComplianceTerms.TextChanged += (s, e) => ResizeComplianceHeaderTextBox(txtComplianceTerms, 46, 180);
            txtComplianceTerms.SizeChanged += (s, e) => ResizeComplianceHeaderTextBox(txtComplianceTerms, 46, 180);

            var titleLayout = new TableLayoutPanel
            {
                AutoSize = true,
                ColumnCount = 2,
                Dock = DockStyle.Fill,
                Margin = new Padding(0, 0, 0, 6),
                RowCount = 1
            };
            titleLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            titleLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 124F));

            txtComplianceTitle = new TextBox
            {
                AcceptsReturn = true,
                AutoSize = false,
                Dock = DockStyle.Fill,
                Font = new Font("Segoe UI", 14F, FontStyle.Bold),
                Height = 72,
                Margin = new Padding(0, 0, 6, 0),
                Multiline = true,
                PlaceholderText = "Compliance plan title",
                TextAlign = HorizontalAlignment.Center,
                WordWrap = true
            };
            txtComplianceTitle.TextChanged += (s, e) => ResizeComplianceHeaderTextBox(txtComplianceTitle, 72, 140);
            txtComplianceTitle.SizeChanged += (s, e) => ResizeComplianceHeaderTextBox(txtComplianceTitle, 72, 140);

            var logoPanel = new Panel
            {
                BorderStyle = BorderStyle.FixedSingle,
                Cursor = Cursors.Hand,
                Dock = DockStyle.Fill,
                Height = 72,
                Margin = new Padding(0),
                MinimumSize = new Size(118, 72)
            };
            picComplianceLogo = new PictureBox
            {
                Cursor = Cursors.Hand,
                Dock = DockStyle.Fill,
                SizeMode = PictureBoxSizeMode.Zoom
            };
            lblComplianceLogoPlaceholder = new Label
            {
                Cursor = Cursors.Hand,
                Dock = DockStyle.Fill,
                Text = "Click to add logo",
                TextAlign = ContentAlignment.MiddleCenter
            };
            logoPanel.Click += ComplianceLogo_Click;
            picComplianceLogo.Click += ComplianceLogo_Click;
            lblComplianceLogoPlaceholder.Click += ComplianceLogo_Click;
            logoPanel.Controls.Add(picComplianceLogo);
            logoPanel.Controls.Add(lblComplianceLogoPlaceholder);
            lblComplianceLogoPlaceholder.BringToFront();

            titleLayout.Controls.Add(txtComplianceTitle, 0, 0);
            titleLayout.Controls.Add(logoPanel, 1, 0);

            var dateTimeLayout = new FlowLayoutPanel
            {
                AutoSize = true,
                Dock = DockStyle.Fill,
                Margin = new Padding(0, 0, 0, 6),
                WrapContents = false
            };
            dateTimeLayout.Controls.Add(new Label
            {
                AutoSize = true,
                Margin = new Padding(0, 7, 8, 0),
                Text = "Date and time:"
            });
            dtpComplianceDateTime = new DateTimePicker
            {
                CustomFormat = "dddd, dd MMMM yyyy",
                Format = DateTimePickerFormat.Custom,
                Margin = new Padding(0),
                Width = 240
            };
            dtpComplianceDateTime.ValueChanged += (s, e) => ResizeComplianceDateTimePicker();
            dtpComplianceDateTime.FontChanged += (s, e) => ResizeComplianceDateTimePicker();
            dateTimeLayout.Controls.Add(dtpComplianceDateTime);
            dateTimeLayout.Controls.Add(new Label
            {
                AutoSize = true,
                Margin = new Padding(14, 7, 8, 0),
                Text = "Time:"
            });
            dtpComplianceTime = new DateTimePicker
            {
                CustomFormat = "HH:mm",
                Format = DateTimePickerFormat.Custom,
                Margin = new Padding(0),
                ShowUpDown = true,
                Width = 82
            };
            dtpComplianceTime.ValueChanged += (s, e) => ResizeComplianceDateTimePicker();
            dtpComplianceTime.FontChanged += (s, e) => ResizeComplianceDateTimePicker();
            dateTimeLayout.Controls.Add(dtpComplianceTime);
            ResizeComplianceDateTimePicker();

            txtComplianceLegend = new TextBox
            {
                AcceptsReturn = true,
                AutoSize = false,
                Dock = DockStyle.Fill,
                Height = 46,
                Margin = new Padding(0, 0, 0, 6),
                Multiline = true,
                PlaceholderText = "Legend / code explanations",
                WordWrap = true
            };
            txtComplianceLegend.TextChanged += (s, e) => ResizeComplianceHeaderTextBox(txtComplianceLegend, 46, 180);
            txtComplianceLegend.SizeChanged += (s, e) => ResizeComplianceHeaderTextBox(txtComplianceLegend, 46, 180);

            var toolbar = new FlowLayoutPanel
            {
                AutoSize = true,
                Dock = DockStyle.Fill,
                Margin = new Padding(0, 0, 0, 8),
                WrapContents = true
            };
            toolbar.Controls.Add(CreateActionButton("New Plan", btnComplianceNew_Click, 96));
            toolbar.Controls.Add(CreateActionButton("Open Plan", btnComplianceOpen_Click, 104));
            toolbar.Controls.Add(CreateActionButton("Save Plan", btnComplianceSave_Click, 100));
            toolbar.Controls.Add(CreateActionButton("Print Plan", btnCompliancePrint_Click, 100));
            toolbar.Controls.Add(CreateActionButton("Add Row", btnComplianceAddRow_Click, 92));
            toolbar.Controls.Add(CreateActionButton("Add Column", btnComplianceAddColumn_Click, 116));
            toolbar.Controls.Add(CreateActionButton("Move Row Up", btnComplianceMoveRowUp_Click, 116));
            toolbar.Controls.Add(CreateActionButton("Move Row Down", btnComplianceMoveRowDown_Click, 132));
            toolbar.Controls.Add(CreateActionButton("Clear Cells", btnComplianceClearCells_Click, 104));
            toolbar.Controls.Add(CreateActionButton("Delete Rows", btnComplianceDeleteRows_Click, 112));
            toolbar.Controls.Add(CreateActionButton("Delete Columns", btnComplianceDeleteColumns_Click, 132));

            var footerLayout = new TableLayoutPanel
            {
                AutoSize = true,
                ColumnCount = 1,
                Dock = DockStyle.Fill,
                Margin = new Padding(0, 8, 0, 0),
                Padding = new Padding(0, 4, 0, 4),
                RowCount = 3
            };
            footerLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            footerLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            footerLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            var signOffLayout = new FlowLayoutPanel
            {
                AutoSize = true,
                Dock = DockStyle.Fill,
                Margin = new Padding(0, 0, 0, 8),
                WrapContents = true
            };
            signOffLayout.Controls.Add(new Label
            {
                AutoSize = true,
                Font = new Font("Segoe UI", 10F),
                Margin = new Padding(0, 0, 36, 0),
                Text = "Completed by: __________"
            });
            signOffLayout.Controls.Add(new Label
            {
                AutoSize = true,
                Font = new Font("Segoe UI", 10F),
                Margin = new Padding(0),
                Text = "Signature: __________"
            });

            var detailIssuesLabel = new Label
            {
                AutoSize = true,
                Font = new Font("Segoe UI", 10F, FontStyle.Bold),
                Margin = new Padding(0, 0, 0, 4),
                Text = "Detail issues with corrective actions"
            };
            txtComplianceDetailIssues = new TextBox
            {
                AcceptsReturn = true,
                AutoSize = false,
                Dock = DockStyle.Fill,
                Height = 108,
                Margin = new Padding(0),
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                WordWrap = true
            };
            footerLayout.Controls.Add(signOffLayout, 0, 0);
            footerLayout.Controls.Add(detailIssuesLabel, 0, 1);
            footerLayout.Controls.Add(txtComplianceDetailIssues, 0, 2);

            dgvCompliancePlan = new DataGridView
            {
                Dock = DockStyle.Fill,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None,
                AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None,
                AllowUserToAddRows = true,
                AllowUserToDeleteRows = true,
                RowHeadersWidth = 52,
                ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize,
                Name = "dgvCompliancePlan"
            };
            dgvCompliancePlan.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dgvCompliancePlan.RowTemplate.MinimumHeight = 28;
            dgvCompliancePlan.CellEndEdit += DgvCompliancePlan_CellEndEdit;
            dgvCompliancePlan.RowPostPaint += DgvCompliancePlan_RowPostPaint;
            ConfigureGrid(dgvCompliancePlan);
            dgvCompliancePlan.ReadOnly = false;
            dgvCompliancePlan.RowHeadersVisible = true;
            dgvCompliancePlan.AllowUserToAddRows = true;
            dgvCompliancePlan.AllowUserToDeleteRows = true;
            dgvCompliancePlan.AllowUserToResizeColumns = true;
            dgvCompliancePlan.AllowUserToResizeRows = true;
            dgvCompliancePlan.MultiSelect = true;
            dgvCompliancePlan.SelectionMode = DataGridViewSelectionMode.CellSelect;

            compliancePlanTable = GridBookEditorService.LoadSheet("", "", 10, 6);
            dgvCompliancePlan.DataSource = compliancePlanTable;
            ResizeCompliancePlanGrid();

            layout.Controls.Add(txtComplianceTerms, 0, 0);
            layout.Controls.Add(titleLayout, 0, 1);
            layout.Controls.Add(dateTimeLayout, 0, 2);
            layout.Controls.Add(txtComplianceLegend, 0, 3);
            layout.Controls.Add(toolbar, 0, 4);
            layout.Controls.Add(dgvCompliancePlan, 0, 5);
            layout.Controls.Add(footerLayout, 0, 6);
            tabCompliance.Controls.Add(layout);
            tabControl1.Controls.Add(tabCompliance);
        }

        private void CreateProcessRecordTab()
        {
            var tabProcessRecord = new TabPage
            {
                Name = "tabProcessRecord",
                Text = "Process Record",
                Padding = new Padding(3)
            };
            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 9,
                Margin = new Padding(0)
            };

            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 50F));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 50F));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            txtProcessRecordVersion = new TextBox
            {
                AcceptsReturn = true,
                AutoSize = false,
                Dock = DockStyle.Fill,
                Height = 46,
                Margin = new Padding(0, 0, 0, 6),
                Multiline = true,
                PlaceholderText = "Version",
                WordWrap = true
            };
            txtProcessRecordVersion.TextChanged += (s, e) => ResizeComplianceHeaderTextBox(txtProcessRecordVersion, 46, 180);
            txtProcessRecordVersion.SizeChanged += (s, e) => ResizeComplianceHeaderTextBox(txtProcessRecordVersion, 46, 180);


            var titleLayout = new TableLayoutPanel
            {
                AutoSize = true,
                ColumnCount = 2,
                Dock = DockStyle.Fill,
                Margin = new Padding(0, 0, 0, 6),
                RowCount = 1
            };
            titleLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            titleLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 124F));

            txtProcessRecordTitle = new TextBox
            {
                AcceptsReturn = true,
                AutoSize = false,
                Dock = DockStyle.Fill,
                Font = new Font("Segoe UI", 14F, FontStyle.Bold),
                Height = 72,
                Margin = new Padding(0, 0, 6, 0),
                Multiline = true,
                PlaceholderText = "Process record title",
                TextAlign = HorizontalAlignment.Center,
                WordWrap = true
            };
            txtProcessRecordTitle.TextChanged += (s, e) => ResizeComplianceHeaderTextBox(txtProcessRecordTitle, 72, 140);
            txtProcessRecordTitle.SizeChanged += (s, e) => ResizeComplianceHeaderTextBox(txtProcessRecordTitle, 72, 140);

            var logoPanel = new Panel
            {
                BorderStyle = BorderStyle.FixedSingle,
                Cursor = Cursors.Hand,
                Dock = DockStyle.Fill,
                Height = 72,
                Margin = new Padding(0),
                MinimumSize = new Size(118, 72)
            };
            picProcessRecordLogo = new PictureBox
            {
                Cursor = Cursors.Hand,
                Dock = DockStyle.Fill,
                SizeMode = PictureBoxSizeMode.Zoom
            };
            lblProcessRecordLogoPlaceholder = new Label
            {
                Cursor = Cursors.Hand,
                Dock = DockStyle.Fill,
                Text = "Click to add logo",
                TextAlign = ContentAlignment.MiddleCenter
            };
            logoPanel.Click += ProcessRecordLogo_Click;
            picProcessRecordLogo.Click += ProcessRecordLogo_Click;
            lblProcessRecordLogoPlaceholder.Click += ProcessRecordLogo_Click;
            logoPanel.Controls.Add(picProcessRecordLogo);
            logoPanel.Controls.Add(lblProcessRecordLogoPlaceholder);
            lblProcessRecordLogoPlaceholder.BringToFront();

            titleLayout.Controls.Add(txtProcessRecordTitle, 0, 0);
            titleLayout.Controls.Add(logoPanel, 1, 0);

            var dateTimeLayout = new FlowLayoutPanel
            {
                AutoSize = true,
                Dock = DockStyle.Fill,
                Margin = new Padding(0, 0, 0, 6),
                WrapContents = false
            };
            dateTimeLayout.Controls.Add(new Label
            {
                AutoSize = true,
                Margin = new Padding(0, 7, 8, 0),
                Text = "Processing Date / Time:"
            });
            dtpProcessRecordDateTime = new DateTimePicker
            {
                CustomFormat = "dddd, dd MMMM yyyy",
                Format = DateTimePickerFormat.Custom,
                Margin = new Padding(0),
                Width = 240
            };
            dtpProcessRecordDateTime.ValueChanged += (s, e) => ResizeProcessRecordDateTimePicker();
            dtpProcessRecordDateTime.FontChanged += (s, e) => ResizeProcessRecordDateTimePicker();
            dateTimeLayout.Controls.Add(dtpProcessRecordDateTime);
            dateTimeLayout.Controls.Add(new Label
            {
                AutoSize = true,
                Margin = new Padding(14, 7, 8, 0),
                Text = "Time:"
            });
            dtpProcessRecordTime = new DateTimePicker
            {
                CustomFormat = "HH:mm",
                Format = DateTimePickerFormat.Custom,
                Margin = new Padding(0),
                ShowUpDown = true,
                Width = 82
            };
            dtpProcessRecordTime.ValueChanged += (s, e) => ResizeProcessRecordDateTimePicker();
            dtpProcessRecordTime.FontChanged += (s, e) => ResizeProcessRecordDateTimePicker();
            dateTimeLayout.Controls.Add(dtpProcessRecordTime);
            ResizeProcessRecordDateTimePicker();

            var supplierNameFarmLayout = new FlowLayoutPanel
            {
                AutoSize = true,
                Dock = DockStyle.Fill,
                Margin = new Padding(0, 0, 0, 8),
                WrapContents = true
            };
            supplierNameFarmLayout.Controls.Add(new Label
            {
                AutoSize = true,
                Font = new Font("Segoe UI", 10F),
                Margin = new Padding(0, 7, 8, 0),
                Text = "Supplier / Farm Name:"
            });
            cboProcessSupplier = new GlossyComboBox
            {
                Margin = new Padding(0, 0, 8, 0),
                DropDownStyle = ComboBoxStyle.DropDown,
                Width = 240
            };
            supplierNameFarmLayout.Controls.Add(cboProcessSupplier);
            supplierNameFarmLayout.Controls.Add(CreateActionButton("Add Supplier", btnProcessAddSupplier_Click, 108));
            supplierNameFarmLayout.Controls.Add(new Label
            {
                AutoSize = true,
                Font = new Font("Segoe UI", 10F),
                Margin = new Padding(20, 7, 8, 0),
                Text = "No. of Birds Killed / Processed:"
            });
            numProcessBirds = new NumericUpDown
            {
                Maximum = 100000000,
                Minimum = 0,
                ThousandsSeparator = true,
                Width = 140
            };
            supplierNameFarmLayout.Controls.Add(numProcessBirds);
            var toolbar = new FlowLayoutPanel
            {
                AutoSize = true,
                Dock = DockStyle.Fill,
                Margin = new Padding(0, 0, 0, 8),
                WrapContents = true
            };
            toolbar.Controls.Add(CreateActionButton("New Plan", btnProcessRecordNew_Click, 96));
            toolbar.Controls.Add(CreateActionButton("Open Plan", btnProcessRecordOpen_Click, 104));
            toolbar.Controls.Add(CreateActionButton("Save Plan", btnProcessRecordSave_Click, 100));
            toolbar.Controls.Add(CreateActionButton("Print Plan", btnProcessRecordPrint_Click, 100));
            cboProcessDropdownLists = new GlossyComboBox
            {
                Margin = new Padding(0, 0, 8, 6),
                Width = 180
            };
            toolbar.Controls.Add(cboProcessDropdownLists);
            toolbar.Controls.Add(CreateActionButton("Create / Edit List", btnProcessDefineDropdown_Click, 132));
            toolbar.Controls.Add(CreateActionButton("Apply Dropdown", btnProcessApplyDropdown_Click, 124));
            toolbar.Controls.Add(CreateActionButton("Remove Dropdown", btnProcessRemoveDropdown_Click, 140));
            toolbar.Controls.Add(CreateActionButton("Delete List", btnProcessDeleteDropdownList_Click, 104));
            toolbar.Controls.Add(CreateActionButton("Add Row", btnProcessRecordAddRow_Click, 92));
            toolbar.Controls.Add(CreateActionButton("Add Column", btnProcessRecordAddColumn_Click, 116));
            toolbar.Controls.Add(CreateActionButton("Rename Column", btnProcessRecordRenameColumn_Click, 128));
            toolbar.Controls.Add(CreateActionButton("Move Row Up", btnProcessRecordMoveRowUp_Click, 116));
            toolbar.Controls.Add(CreateActionButton("Move Row Down", btnProcessRecordMoveRowDown_Click, 132));
            toolbar.Controls.Add(CreateActionButton("Clear Cells", btnProcessRecordClearCells_Click, 104));
            toolbar.Controls.Add(CreateActionButton("Delete Rows", btnProcessRecordDeleteRows_Click, 112));
            toolbar.Controls.Add(CreateActionButton("Delete Columns", btnProcessRecordDeleteColumns_Click, 132));

            var middleProcessRecordPanel = new FlowLayoutPanel
            {
                AutoSize = true,
                Dock = DockStyle.Fill,
                Margin = new Padding(0, 8, 0, 8)
            };

            middleProcessRecordPanel.Controls.Add(new Label 
            {
                AutoSize = true,
                Font = new Font("Segoe UI", 10F),
                Margin = new Padding(0, 7, 8, 0),
                Text = "Inwards Number:" 
            });

            var cboInwardsNumber = new GlossyComboBox
            {
                Margin = new Padding(0, 0, 8, 0),
                DropDownStyle = ComboBoxStyle.DropDown,
                Width = 240
            };
            middleProcessRecordPanel.Controls.Add(cboInwardsNumber);

            middleProcessRecordPanel.Controls.Add(new Label
            {
                AutoSize = true,
                Font = new Font("Segoe UI", 10F),
                Margin = new Padding(0, 7, 8, 0),
                Text = "QC Checked by:"
            });

            var cboQualityControl = new GlossyComboBox
            {
                Margin = new Padding(0, 0, 8, 0),
                DropDownStyle = ComboBoxStyle.DropDown,
                Width = 300
            };
            middleProcessRecordPanel.Controls.Add(cboQualityControl);

            var footerLayout = new TableLayoutPanel
            {
                AutoSize = true,
                ColumnCount = 1,
                Dock = DockStyle.Fill,
                Margin = new Padding(0, 8, 0, 0),
                Padding = new Padding(0, 4, 0, 4),
                RowCount = 3
            };
            footerLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            footerLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            footerLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            var signOffLayout = new FlowLayoutPanel
            {
                AutoSize = true,
                Dock = DockStyle.Fill,
                Margin = new Padding(0, 0, 0, 8),
                WrapContents = true
            };
            signOffLayout.Controls.Add(new Label
            {
                AutoSize = true,
                Font = new Font("Segoe UI", 10F),
                Margin = new Padding(0, 0, 36, 0),
                Text = "Completed by: __________"
            });
            signOffLayout.Controls.Add(new Label
            {
                AutoSize = true,
                Font = new Font("Segoe UI", 10F),
                Margin = new Padding(0),
                Text = "Signature: __________"
            });
            footerLayout.Controls.Add(signOffLayout, 0, 0);

            dgvDailyProduction = new DataGridView
            {
                Dock = DockStyle.Fill,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None,
                AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None,
                AllowUserToAddRows = true,
                AllowUserToDeleteRows = true,
                RowHeadersWidth = 52,
                ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize,
                Name = "dgvDailyProduction"
            };
            dgvDailyProduction.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dgvDailyProduction.RowTemplate.MinimumHeight = 28;
            dgvDailyProduction.CellEndEdit += DgvDailyProduction_CellEndEdit;
            dgvDailyProduction.RowPostPaint += DgvDailyProduction_RowPostPaint;
            ConfigureGrid(dgvDailyProduction);
            dgvDailyProduction.ReadOnly = false;
            dgvDailyProduction.RowHeadersVisible = true;
            dgvDailyProduction.AllowUserToAddRows = true;
            dgvDailyProduction.AllowUserToDeleteRows = true;
            dgvDailyProduction.AllowUserToResizeColumns = true;
            dgvDailyProduction.AllowUserToResizeRows = true;
            dgvDailyProduction.MultiSelect = true;
            dgvDailyProduction.SelectionMode = DataGridViewSelectionMode.CellSelect;

            dailyProductionTable = GridBookEditorService.LoadSheet("", "", 10, 6);
            dgvDailyProduction.DataSource = dailyProductionTable;
            ResizeDailyProductionGrid();

            dgvProcessRecord = new DataGridView
            {
                Dock = DockStyle.Fill,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None,
                AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None,
                AllowUserToAddRows = true,
                AllowUserToDeleteRows = true,
                RowHeadersWidth = 52,
                ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize,
                Name = "dgvProcessRecord"
            };
            dgvProcessRecord.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dgvProcessRecord.RowTemplate.MinimumHeight = 28;
            dgvProcessRecord.CellEndEdit += DgvProcessRecord_CellEndEdit;
            dgvProcessRecord.RowPostPaint += DgvProcessRecord_RowPostPaint;
            ConfigureGrid(dgvProcessRecord);
            dgvProcessRecord.ReadOnly = false;
            dgvProcessRecord.RowHeadersVisible = true;
            dgvProcessRecord.AllowUserToAddRows = true;
            dgvProcessRecord.AllowUserToDeleteRows = true;
            dgvProcessRecord.AllowUserToResizeColumns = true;
            dgvProcessRecord.AllowUserToResizeRows = true;
            dgvProcessRecord.MultiSelect = true;
            dgvProcessRecord.SelectionMode = DataGridViewSelectionMode.CellSelect;

            processRecordTable = GridBookEditorService.LoadSheet("", "", 10, 6);
            dgvProcessRecord.DataSource = processRecordTable;
            ResizeProcessRecordGrid();

            layout.Controls.Add(txtProcessRecordVersion, 0, 0);
            layout.Controls.Add(titleLayout, 0, 1);
            layout.Controls.Add(dateTimeLayout, 0, 2);
            layout.Controls.Add(supplierNameFarmLayout, 0, 3);
            layout.Controls.Add(toolbar, 0, 4);
            layout.Controls.Add(dgvProcessRecord, 0, 5);
            layout.Controls.Add(middleProcessRecordPanel, 0, 6);
            layout.Controls.Add(dgvDailyProduction, 0, 7);
            layout.Controls.Add(footerLayout, 0, 8);
            tabProcessRecord.Controls.Add(layout);
            tabControl1.Controls.Add(tabProcessRecord);
        }
        private void ResizeComplianceHeaderTextBox(TextBox textBox, int minimumHeight, int maximumHeight)
        {
            DocumentEditorService.ResizeTextBox(textBox, minimumHeight, maximumHeight);
        }
        private void ResizeComplianceDateTimePicker()
        {
            DocumentEditorService.ResizeDateTimePickers(dtpComplianceDateTime, dtpComplianceTime);
        }

        private void ResizeProcessRecordDateTimePicker()
        {
            DocumentEditorService.ResizeDateTimePickers(dtpProcessRecordDateTime, dtpProcessRecordTime);
        }
        private void ComplianceLogo_Click(object sender, EventArgs e)
        {
            DocumentEditorService.SelectLogo(this, picComplianceLogo, lblComplianceLogoPlaceholder);
        }

        private void ProcessRecordLogo_Click(object sender, EventArgs e)
        {
            DocumentEditorService.SelectLogo(this, picProcessRecordLogo, lblProcessRecordLogoPlaceholder);
        }
        private void ResizeCompliancePlanGrid()
        {
            DocumentEditorService.ResizeGrid(dgvCompliancePlan);
        }

        private void ResizeProcessRecordGrid()
        {
            DocumentEditorService.ResizeGrid(dgvProcessRecord);
        }
        private void ResizeDailyProductionGrid()
        {
            DocumentEditorService.ResizeGrid(dgvDailyProduction);
        }
        private void DgvCompliancePlan_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            ResizeCompliancePlanGrid();
        }

        private void DgvProcessRecord_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            ResizeProcessRecordGrid();
        }
        private void DgvDailyProduction_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            ResizeProcessRecordGrid();
        }
        private void DgvCompliancePlan_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            DocumentEditorService.DrawRowNumber(dgvCompliancePlan, e, Color.FromArgb(0, 255, 40));
        }

        private void DgvProcessRecord_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            DocumentEditorService.DrawRowNumber(dgvProcessRecord, e, Color.FromArgb(0, 255, 40));
        }
        private void DgvDailyProduction_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            DocumentEditorService.DrawRowNumber(dgvDailyProduction, e, Color.FromArgb(0, 255, 40));
        }
        private void btnComplianceNew_Click(object sender, EventArgs e)
        {
            compliancePlanPath = "";
            txtComplianceTerms.Clear();
            txtComplianceTitle.Clear();
            txtComplianceLegend.Clear();
            txtComplianceDetailIssues.Clear();
            dtpComplianceDateTime.Value = DateTime.Now;
            dtpComplianceTime.Value = DateTime.Now;
            DocumentEditorService.ResetLogo(picComplianceLogo, lblComplianceLogoPlaceholder);
            compliancePlanTable = GridBookEditorService.LoadSheet("", "", 10, 6);
            dgvCompliancePlan.DataSource = compliancePlanTable;
            ResizeCompliancePlanGrid();
        }

        private void RefreshProcessSupplierChoices()
        {
            if (cboProcessSupplier == null || appSettings == null)
                return;

            appSettings.SupplierFarmNames ??= new List<string>();
            string selected = processRecordMetadata.SupplierFarmName;
            if (string.IsNullOrWhiteSpace(selected))
                selected = cboProcessSupplier.SelectedItem?.ToString() ?? "";

            List<string> suppliers = appSettings.SupplierFarmNames
                .Where(name => !string.IsNullOrWhiteSpace(name))
                .Select(name => name.Trim())
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(name => name)
                .ToList();
            if (!string.IsNullOrWhiteSpace(selected) &&
                !suppliers.Contains(selected, StringComparer.OrdinalIgnoreCase))
            {
                suppliers.Add(selected);
            }

            cboProcessSupplier.Items.Clear();
            cboProcessSupplier.Items.AddRange(suppliers.Cast<object>().ToArray());
            cboProcessSupplier.SelectedItem = suppliers.FirstOrDefault(name =>
                string.Equals(name, selected, StringComparison.OrdinalIgnoreCase));
            if (cboProcessSupplier.SelectedIndex < 0 && cboProcessSupplier.Items.Count > 0)
                cboProcessSupplier.SelectedIndex = 0;
        }

        private void RefreshProcessDropdownLists(string selectedName = null)
        {
            if (cboProcessDropdownLists == null)
                return;

            selectedName ??= cboProcessDropdownLists.SelectedItem?.ToString();
            string[] names = processRecordMetadata.DropdownLists.Keys
                .OrderBy(name => name)
                .ToArray();
            cboProcessDropdownLists.Items.Clear();
            cboProcessDropdownLists.Items.AddRange(names.Cast<object>().ToArray());
            cboProcessDropdownLists.SelectedItem = names.FirstOrDefault(name =>
                string.Equals(name, selectedName, StringComparison.OrdinalIgnoreCase));
            if (cboProcessDropdownLists.SelectedIndex < 0 && names.Length > 0)
                cboProcessDropdownLists.SelectedIndex = 0;
        }

        private void btnProcessAddSupplier_Click(object sender, EventArgs e)
        {
            if (!TextPromptDialog.Show(this, "Add Supplier / Farm", "Supplier or farm name:", "", false, out string name) ||
                string.IsNullOrWhiteSpace(name))
            {
                return;
            }

            appSettings.SupplierFarmNames ??= new List<string>();
            if (!appSettings.SupplierFarmNames.Contains(name, StringComparer.OrdinalIgnoreCase))
            {
                appSettings.SupplierFarmNames.Add(name);
                appSettings.SupplierFarmNames = appSettings.SupplierFarmNames
                    .OrderBy(item => item)
                    .ToList();
                SettingsManager.Save(appSettings);
            }

            processRecordMetadata.SupplierFarmName = name;
            RefreshProcessSupplierChoices();
        }

        private void btnProcessDefineDropdown_Click(object sender, EventArgs e)
        {
            string currentName = cboProcessDropdownLists.SelectedItem?.ToString() ?? "New List";
            if (!TextPromptDialog.Show(this, "Dropdown List", "List name:", currentName, false, out string listName) ||
                string.IsNullOrWhiteSpace(listName))
            {
                return;
            }

            string existingItems = processRecordMetadata.DropdownLists.TryGetValue(listName, out List<string> existing)
                ? string.Join(Environment.NewLine, existing)
                : "";
            if (!TextPromptDialog.Show(
                    this,
                    "Populate Dropdown List",
                    "Enter one item per line:",
                    existingItems,
                    true,
                    out string itemText))
            {
                return;
            }

            List<string> items = itemText
                .Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
            if (items.Count == 0)
            {
                MessageBox.Show(this, "Add at least one dropdown item.", "Dropdown List", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            processRecordMetadata.DropdownLists[listName] = items;
            ProcessRecordGridService.UpdateListCells(dgvProcessRecord, listName, items);
            RefreshProcessDropdownLists(listName);
        }

        private void btnProcessApplyDropdown_Click(object sender, EventArgs e)
        {
            string listName = cboProcessDropdownLists.SelectedItem?.ToString();
            if (string.IsNullOrWhiteSpace(listName) ||
                !processRecordMetadata.DropdownLists.TryGetValue(listName, out List<string> items))
            {
                MessageBox.Show(this, "Create or select a dropdown list first.", "Apply Dropdown", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (dgvProcessRecord.SelectedCells.Count == 0)
            {
                MessageBox.Show(this, "Select one or more grid cells first.", "Apply Dropdown", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            ProcessRecordGridService.AssignSelectedCells(dgvProcessRecord, listName, items);
        }

        private void btnProcessRemoveDropdown_Click(object sender, EventArgs e)
        {
            ProcessRecordGridService.RemoveDropdownsFromSelectedCells(dgvProcessRecord);
        }

        private void btnProcessDeleteDropdownList_Click(object sender, EventArgs e)
        {
            string listName = cboProcessDropdownLists.SelectedItem?.ToString();
            if (string.IsNullOrWhiteSpace(listName))
                return;

            if (MessageBox.Show(
                    this,
                    $"Delete the dropdown list '{listName}' and remove it from all cells?",
                    "Delete Dropdown List",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning) != DialogResult.Yes)
            {
                return;
            }

            ProcessRecordGridService.RemoveListCells(dgvProcessRecord, listName);
            processRecordMetadata.DropdownLists.Remove(listName);
            RefreshProcessDropdownLists();
        }

        private void btnProcessRecordNew_Click(object sender, EventArgs e)
        {
            processRecordPath = "";
            processRecordMetadata = new ProcessRecordMetadata();
            txtProcessRecordVersion.Clear();
            txtProcessRecordTitle.Clear();
            numProcessBirds.Value = 0;
            dtpProcessRecordDateTime.Value = DateTime.Now;
            dtpProcessRecordTime.Value = DateTime.Now;
            DocumentEditorService.ResetLogo(picProcessRecordLogo, lblProcessRecordLogoPlaceholder);
            processRecordTable = GridBookEditorService.LoadSheet("", "", 10, 6);
            dgvProcessRecord.DataSource = processRecordTable;
            RefreshProcessSupplierChoices();
            RefreshProcessDropdownLists();
            ResizeProcessRecordGrid();
        }
        private void btnComplianceOpen_Click(object sender, EventArgs e)
        {
            using var dialog = new OpenFileDialog
            {
                Filter = "Excel Files (*.xlsx)|*.xlsx"
            };

            if (dialog.ShowDialog(this) != DialogResult.OK)
                return;

            string[] sheets = GridBookEditorService.GetSheetNames(dialog.FileName);
            if (sheets.Length == 0)
                return;

            string sheetName = sheets.FirstOrDefault(name =>
                string.Equals(name, "Compliance Plan", StringComparison.OrdinalIgnoreCase)) ?? sheets[0];

            compliancePlanPath = dialog.FileName;
            txtComplianceTerms.Clear();
            txtComplianceTitle.Text = Path.GetFileNameWithoutExtension(dialog.FileName);
            txtComplianceLegend.Clear();
            txtComplianceDetailIssues.Clear();
            dtpComplianceDateTime.Value = DateTime.Now;
            dtpComplianceTime.Value = DateTime.Now;
            DocumentEditorService.ResetLogo(picComplianceLogo, lblComplianceLogoPlaceholder);
            compliancePlanTable = GridBookEditorService.LoadSheet(compliancePlanPath, sheetName, 10, 6);
            dgvCompliancePlan.DataSource = compliancePlanTable;
            ResizeCompliancePlanGrid();
        }

        private void btnProcessRecordOpen_Click(object sender, EventArgs e)
        {
            using var dialog = new OpenFileDialog
            {
                Filter = "Excel Files (*.xlsx)|*.xlsx"
            };

            if (dialog.ShowDialog(this) != DialogResult.OK)
                return;

            string[] sheets = GridBookEditorService.GetSheetNames(dialog.FileName);
            if (sheets.Length == 0)
                return;

            string sheetName = sheets.FirstOrDefault(name =>
                string.Equals(name, "Process Record", StringComparison.OrdinalIgnoreCase)) ?? sheets[0];

            processRecordPath = dialog.FileName;
            processRecordMetadata = ProcessRecordGridService.LoadMetadata(processRecordPath);
            txtProcessRecordVersion.Clear();
            txtProcessRecordTitle.Text = Path.GetFileNameWithoutExtension(dialog.FileName);
            numProcessBirds.Value = Math.Clamp(
                processRecordMetadata.BirdsProcessed,
                (int)numProcessBirds.Minimum,
                (int)numProcessBirds.Maximum);
            dtpProcessRecordDateTime.Value = DateTime.Now;
            dtpProcessRecordTime.Value = DateTime.Now;
            DocumentEditorService.ResetLogo(picProcessRecordLogo, lblProcessRecordLogoPlaceholder);
            processRecordTable = GridBookEditorService.LoadSheet(processRecordPath, sheetName, 10, 6);
            dgvProcessRecord.DataSource = processRecordTable;
            RefreshProcessSupplierChoices();
            RefreshProcessDropdownLists();
            ProcessRecordGridService.ApplyAssignments(dgvProcessRecord, processRecordMetadata);
            ResizeProcessRecordGrid();
        }
        private void btnComplianceSave_Click(object sender, EventArgs e)
        {
            dgvCompliancePlan.EndEdit();

            if (string.IsNullOrWhiteSpace(compliancePlanPath))
            {
                using var dialog = new SaveFileDialog
                {
                    Filter = "Excel Files (*.xlsx)|*.xlsx",
                    FileName = "CompliancePlan.xlsx"
                };

                if (dialog.ShowDialog(this) != DialogResult.OK)
                    return;

                compliancePlanPath = dialog.FileName;
            }

            GridBookEditorService.SaveSheet(compliancePlanPath, "Compliance Plan", compliancePlanTable);
            NotificationManager.ShowNotification("Compliance Plan Saved", compliancePlanPath);
        }

        private void btnProcessRecordSave_Click(object sender, EventArgs e)
        {
            dgvProcessRecord.EndEdit();

            if (string.IsNullOrWhiteSpace(processRecordPath))
            {
                using var dialog = new SaveFileDialog
                {
                    Filter = "Excel Files (*.xlsx)|*.xlsx",
                    FileName = "ProcessRecord.xlsx"
                };

                if (dialog.ShowDialog(this) != DialogResult.OK)
                    return;

                processRecordPath = dialog.FileName;
            }

            processRecordMetadata.SupplierFarmName = cboProcessSupplier.SelectedItem?.ToString() ?? "";
            processRecordMetadata.BirdsProcessed = (int)numProcessBirds.Value;
            processRecordMetadata.CellDropdownAssignments =
                ProcessRecordGridService.CaptureAssignments(dgvProcessRecord);
            GridBookEditorService.SaveSheet(processRecordPath, "Process Record", processRecordTable);
            ProcessRecordGridService.SaveMetadata(processRecordPath, processRecordMetadata);
            NotificationManager.ShowNotification("Process Record Saved", processRecordPath);
        }
        private void btnCompliancePrint_Click(object sender, EventArgs e)
        {
            dgvCompliancePlan.EndEdit();
            CompliancePlanPrintService.Print(
                this,
                txtComplianceTerms.Text,
                txtComplianceTitle.Text,
                dtpComplianceDateTime.Value.Date + dtpComplianceTime.Value.TimeOfDay,
                txtComplianceLegend.Text,
                txtComplianceDetailIssues.Text,
                compliancePlanTable,
                picComplianceLogo.Image);
        }
        private void btnProcessRecordPrint_Click(object sender, EventArgs e)
        {
            dgvProcessRecord.EndEdit();
            CompliancePlanPrintService.PrintProcessRecord(
                this,
                txtProcessRecordVersion.Text,
                txtProcessRecordTitle.Text,
                dtpProcessRecordDateTime.Value.Date + dtpProcessRecordTime.Value.TimeOfDay,
                cboProcessSupplier.SelectedItem?.ToString() ?? "",
                (int)numProcessBirds.Value,
                processRecordTable,
                picProcessRecordLogo.Image);
        }

        private void btnComplianceAddRow_Click(object sender, EventArgs e)
        {
            DocumentEditorService.AddRow(compliancePlanTable, dgvCompliancePlan);
        }

        private void btnComplianceAddColumn_Click(object sender, EventArgs e)
        {
            DocumentEditorService.AddColumn(compliancePlanTable, dgvCompliancePlan);
        }

        private void btnProcessRecordAddRow_Click(object sender, EventArgs e)
        {
            DocumentEditorService.AddRow(processRecordTable, dgvProcessRecord);
        }

        private void btnProcessRecordAddColumn_Click(object sender, EventArgs e)
        {
            DocumentEditorService.AddColumn(processRecordTable, dgvProcessRecord);
        }

        private void btnProcessRecordRenameColumn_Click(object sender, EventArgs e)
        {
            int columnIndex = dgvProcessRecord.CurrentCell?.ColumnIndex ?? -1;
            if (columnIndex < 0 || columnIndex >= processRecordTable.Columns.Count)
            {
                MessageBox.Show(this, "Select a grid column first.", "Rename Column", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            string currentName = processRecordTable.Columns[columnIndex].ColumnName;
            if (!TextPromptDialog.Show(this, "Rename Column", "Column name:", currentName, false, out string newName) ||
                string.IsNullOrWhiteSpace(newName) ||
                string.Equals(currentName, newName, StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            if (processRecordTable.Columns.Contains(newName))
            {
                MessageBox.Show(this, "That column name is already in use.", "Rename Column", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            processRecordTable.Columns[columnIndex].ColumnName = newName;
            ResizeProcessRecordGrid();
        }

        private void btnComplianceMoveRowUp_Click(object sender, EventArgs e)
        {
            MoveSelectedComplianceRows(-1);
        }

        private void btnComplianceMoveRowDown_Click(object sender, EventArgs e)
        {
            MoveSelectedComplianceRows(1);
        }

        private void MoveSelectedComplianceRows(int direction)
        {
            DocumentEditorService.MoveSelectedRows(compliancePlanTable, dgvCompliancePlan, direction);
        }

        private void btnProcessRecordMoveRowUp_Click(object sender, EventArgs e)
        {
            MoveSelectedProcessRecordRows(-1);
        }

        private void btnProcessRecordMoveRowDown_Click(object sender, EventArgs e)
        {
            MoveSelectedProcessRecordRows(1);
        }

        private void MoveSelectedProcessRecordRows(int direction)
        {
            DocumentEditorService.MoveSelectedRows(processRecordTable, dgvProcessRecord, direction);
        }
        private void btnComplianceClearCells_Click(object sender, EventArgs e)
        {
            DocumentEditorService.ClearSelectedCells(dgvCompliancePlan);
        }

        private void btnComplianceDeleteRows_Click(object sender, EventArgs e)
        {
            DocumentEditorService.DeleteSelectedRows(compliancePlanTable, dgvCompliancePlan);
        }

        private void btnComplianceDeleteColumns_Click(object sender, EventArgs e)
        {
            DocumentEditorService.DeleteSelectedColumns(compliancePlanTable, dgvCompliancePlan);
        }

        private void btnProcessRecordClearCells_Click(object sender, EventArgs e)
        {
            DocumentEditorService.ClearSelectedCells(dgvProcessRecord);
        }

        private void btnProcessRecordDeleteRows_Click(object sender, EventArgs e)
        {
            DocumentEditorService.DeleteSelectedRows(processRecordTable, dgvProcessRecord);
        }

        private void btnProcessRecordDeleteColumns_Click(object sender, EventArgs e)
        {
            DocumentEditorService.DeleteSelectedColumns(processRecordTable, dgvProcessRecord);
        }
        private void CreateBehaviorControls()
        {
            btnNewExcel.Visible = false;
            btnSaveExcel.Visible = false;
            btnExportExcel.Visible = false;
            chkMinimizeTray.Visible = false;

            var lblMinimize = new Label
            {
                AutoSize = true,
                Margin = new Padding(0, 8, 4, 0),
                Text = "Minimize"
            };

            cboMinimizeBehavior = new GlossyComboBox
            {
                Width = 120,
                Margin = new Padding(0, 4, 10, 6)
            };
            cboMinimizeBehavior.Items.AddRange(new object[] { "Tray", "Window" });
            cboMinimizeBehavior.SelectedIndex = 0;

            chkReminderAgentOnClose = new GlossyCheckBox
            {
                AutoSize = true,
                Margin = new Padding(0, 6, 10, 6),
                Text = "Reminder agent on close"
            };

            actionLayout.Controls.Add(lblMinimize);
            actionLayout.Controls.Add(cboMinimizeBehavior);
            actionLayout.Controls.Add(chkReminderAgentOnClose);
        }

        private void CreateChartExportControls()
        {
            var chartToolbar = new FlowLayoutPanel
            {
                AutoSize = true,
                Dock = DockStyle.Top,
                Margin = new Padding(0, 0, 0, 8)
            };
            chartToolbar.Controls.Add(CreateActionButton("Export Status Pie", (s, e) => ExportChart(statusPieChart, "StatusPie.png"), 142));
            chartToolbar.Controls.Add(CreateActionButton("Export Open Pie", (s, e) => ExportChart(openPieChart, "OpenTrainingPie.png"), 136));

            tabCharts.Controls.Remove(chartsLayout);

            var chartPageLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 2,
                Margin = new Padding(0)
            };
            chartPageLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            chartPageLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            chartPageLayout.Controls.Add(chartToolbar, 0, 0);
            chartPageLayout.Controls.Add(chartsLayout, 0, 1);
            tabCharts.Controls.Add(chartPageLayout);
        }

        private void TabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            ApplyFocusedTabLayout();
        }

        private void ApplyFocusedTabLayout()
        {
            bool focusedGridMode = tabControl1.SelectedTab?.Name is
                "tabGridBookEditor" or
                "tabCompliancePlan" or
                "tabProcessRecord";

            topLayout.Visible = !focusedGridMode;
            fileSearchLayout.Visible = !focusedGridMode;
            actionLayout.Visible = !focusedGridMode;
            picAccentBar.Visible = !focusedGridMode;
            dashboardLayout.Visible = !focusedGridMode;
            footerLayout.Visible = !focusedGridMode;
            tabControl1.Margin = focusedGridMode
                ? new Padding(8, 8, 8, 8)
                : new Padding(12, 0, 12, 12);
        }

        private GlossyButton CreateActionButton(string text, EventHandler handler, int width)
        {
            var button = new GlossyButton
            {
                AutoSize = true,
                Margin = new Padding(0, 0, 8, 6),
                MinimumSize = new Size(width, 34),
                Text = text
            };
            button.Click += handler;
            return button;
        }

        private void LoadGridBookEditor(string path)
        {
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
                return;

            string previousSheet = cboGridBookSheets.SelectedItem?.ToString();
            cboGridBookSheets.Items.Clear();
            cboGridBookSheets.Items.AddRange(GridBookEditorService.GetSheetNames(path));

            if (cboGridBookSheets.Items.Count == 0)
                return;

            int index = !string.IsNullOrWhiteSpace(previousSheet) && cboGridBookSheets.Items.Contains(previousSheet)
                ? cboGridBookSheets.Items.IndexOf(previousSheet)
                : 0;

            cboGridBookSheets.SelectedIndex = index;
        }

        private void LoadGridBookSheet(string sheetName)
        {
            if (string.IsNullOrWhiteSpace(excelPath) || !File.Exists(excelPath) || string.IsNullOrWhiteSpace(sheetName))
                return;

            currentGridBookSheet = sheetName;
            currentGridBookTable = GridBookEditorService.LoadSheet(excelPath, sheetName);
            dgvGridBook.DataSource = currentGridBookTable;
            foreach (DataGridViewColumn column in dgvGridBook.Columns)
            {
                column.Width = Math.Max(column.Width, 110);
            }

            tabControl1.SelectedTab = tabControl1.TabPages["tabGridBookEditor"];
        }

        private void SaveCurrentGridBookSheet()
        {
            dgvGridBook.EndEdit();

            if (string.IsNullOrWhiteSpace(excelPath))
            {
                using var dialog = new SaveFileDialog
                {
                    Filter = "Excel Files (*.xlsx)|*.xlsx",
                    FileName = "GridBook.xlsx"
                };

                if (dialog.ShowDialog(this) != DialogResult.OK)
                    return;

                excelPath = dialog.FileName;
                lblFile.Text = excelPath;
            }

            GridBookEditorService.SaveSheet(excelPath, currentGridBookSheet, currentGridBookTable);
            SaveSettingsFromControls();
            LoadGridBookEditor(excelPath);
            NotificationManager.ShowNotification("GridBook Saved", $"{currentGridBookSheet} saved.");
        }

        private void btnEditorNew_Click(object sender, EventArgs e)
        {
            using var dialog = new SaveFileDialog
            {
                Filter = "Excel Files (*.xlsx)|*.xlsx",
                FileName = "GridBook.xlsx"
            };

            if (dialog.ShowDialog(this) != DialogResult.OK)
                return;

            GridBookEditorService.CreateBlankGridBook(dialog.FileName);
            excelPath = dialog.FileName;
            lblFile.Text = excelPath;
            SaveSettingsFromControls();
            StartWatcher();
            LoadGridBookEditor(excelPath);
        }

        private void btnEditorOpen_Click(object sender, EventArgs e)
        {
            using var dialog = new OpenFileDialog
            {
                Filter = "Excel Files (*.xlsx)|*.xlsx"
            };

            if (dialog.ShowDialog(this) != DialogResult.OK)
                return;

            excelPath = dialog.FileName;
            lblFile.Text = excelPath;
            SaveSettingsFromControls();
            StartWatcher();
            LoadGridBookEditor(excelPath);
        }

        private void btnEditorSave_Click(object sender, EventArgs e)
        {
            SaveCurrentGridBookSheet();
        }

        private void btnEditorSaveAs_Click(object sender, EventArgs e)
        {
            using var dialog = new SaveFileDialog
            {
                Filter = "Excel Files (*.xlsx)|*.xlsx",
                FileName = Path.GetFileName(excelPath)
            };

            if (dialog.ShowDialog(this) != DialogResult.OK)
                return;

            if (!string.IsNullOrWhiteSpace(excelPath) && File.Exists(excelPath))
            {
                SaveCurrentGridBookSheet();
                GridBookEditorService.ExportGridBook(excelPath, dialog.FileName);
            }
            else
            {
                dgvGridBook.EndEdit();
                GridBookEditorService.SaveSheet(dialog.FileName, currentGridBookSheet, currentGridBookTable);
            }

            excelPath = dialog.FileName;
            lblFile.Text = excelPath;
            SaveSettingsFromControls();
            StartWatcher();
            LoadGridBookEditor(excelPath);
            NotificationManager.ShowNotification("GridBook Saved As", dialog.FileName);
        }

        private void btnEditorExport_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(excelPath) || !File.Exists(excelPath))
                return;

            using var dialog = new SaveFileDialog
            {
                Filter = "Excel Files (*.xlsx)|*.xlsx",
                FileName = Path.GetFileNameWithoutExtension(excelPath) + "-export.xlsx"
            };

            if (dialog.ShowDialog(this) != DialogResult.OK)
                return;

            SaveCurrentGridBookSheet();
            GridBookEditorService.ExportGridBook(excelPath, dialog.FileName);
            NotificationManager.ShowNotification("GridBook Exported", dialog.FileName);
        }

        private void btnEditorAddSheet_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(excelPath))
            {
                btnEditorNew_Click(sender, e);
                if (string.IsNullOrWhiteSpace(excelPath))
                    return;
            }

            string sheetName = $"Sheet{cboGridBookSheets.Items.Count + 1}";
            GridBookEditorService.AddSheet(excelPath, sheetName);
            LoadGridBookEditor(excelPath);
            cboGridBookSheets.SelectedItem = sheetName;
        }

        private void btnEditorAddRow_Click(object sender, EventArgs e)
        {
            if (currentGridBookTable.Columns.Count == 0)
            {
                currentGridBookTable = GridBookEditorService.LoadSheet(excelPath, currentGridBookSheet);
                dgvGridBook.DataSource = currentGridBookTable;
            }

            currentGridBookTable.Rows.Add(currentGridBookTable.NewRow());
        }

        private void btnEditorAddColumn_Click(object sender, EventArgs e)
        {
            string name = GridBookEditorService.ColumnName(currentGridBookTable.Columns.Count);
            currentGridBookTable.Columns.Add(name);
        }

       private void btnEditorTrainingColors_Click(object sender, EventArgs e)
        {
            ThemeManager.ApplyTrainingColors(dgvGridBook);
        }
        private void cboGridBookSheets_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cboGridBookSheets.SelectedItem == null)
                return;

            LoadGridBookSheet(cboGridBookSheets.SelectedItem.ToString());
        }

        private void DgvGridBook_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            string rowNumber = (e.RowIndex + 1).ToString();
            TextRenderer.DrawText(
                e.Graphics,
                rowNumber,
                dgvGridBook.Font,
                new Rectangle(e.RowBounds.Left, e.RowBounds.Top, dgvGridBook.RowHeadersWidth - 4, e.RowBounds.Height),
                Color.FromArgb(0, 255, 40),
                TextFormatFlags.Right | TextFormatFlags.VerticalCenter);
        }

        private void ExportChart(Control chart, string defaultName)
        {
            using var dialog = new SaveFileDialog
            {
                Filter = "PNG Image (*.png)|*.png",
                FileName = defaultName
            };

            if (dialog.ShowDialog(this) != DialogResult.OK)
                return;

            ChartExportService.Export(chart, dialog.FileName);
            NotificationManager.ShowNotification("Chart Exported", dialog.FileName);
        }

        private void btnNewExcel_Click(object sender, EventArgs e)
        {
            using var dialog = new SaveFileDialog
            {
                Filter = "Excel Files (*.xlsx)|*.xlsx",
                FileName = "TrainingMonitor.xlsx"
            };

            if (dialog.ShowDialog(this) != DialogResult.OK)
                return;

            ExcelMonitor.CreateTemplate(dialog.FileName);
            excelPath = dialog.FileName;
            lblFile.Text = excelPath;
            SaveSettingsFromControls();
            StartWatcher();
            RunScan();
            LoadGridBookEditor(excelPath);
            NotificationManager.ShowNotification("Excel Created", dialog.FileName);
        }

        private void btnSaveExcel_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(excelPath))
            {
                using var dialog = new SaveFileDialog
                {
                    Filter = "Excel Files (*.xlsx)|*.xlsx",
                    FileName = "TrainingMonitor.xlsx"
                };

                if (dialog.ShowDialog(this) != DialogResult.OK)
                    return;

                excelPath = dialog.FileName;
                lblFile.Text = excelPath;
            }

            List<TrainingAlert> alerts = GetEditableAlerts();
            ExcelMonitor.SaveTrainingData(excelPath, alerts);
            currentAlerts = alerts;
            SaveSettingsFromControls();
            UpdateDashboard(alerts);
            UpdateCharts(alerts);
            LoadGridBookEditor(excelPath);
            NotificationManager.ShowNotification("Excel Saved", "Training grid edits were saved.");
        }

        private void btnExportExcel_Click(object sender, EventArgs e)
        {
            using var dialog = new SaveFileDialog
            {
                Filter = "Excel Files (*.xlsx)|*.xlsx",
                FileName = "TrainingExport.xlsx"
            };

            if (dialog.ShowDialog(this) != DialogResult.OK)
                return;

            ExcelMonitor.ExportFlatWorkbook(dialog.FileName, GetEditableAlerts());
            NotificationManager.ShowNotification("Excel Exported", dialog.FileName);
        }

        private void btnExportCsv_Click(object sender, EventArgs e)
        {
            bool exportHistory = tabControl1.SelectedTab == tabHistory;
            using var dialog = new SaveFileDialog
            {
                Filter = "CSV Files (*.csv)|*.csv",
                FileName = exportHistory ? "TrainingHistory.csv" : "TrainingAlerts.csv"
            };

            if (dialog.ShowDialog(this) != DialogResult.OK)
                return;

            if (exportHistory)
            {
                ExportService.ExportHistoryCsv(dialog.FileName, HistoryManager.GetHistory());
            }
            else
            {
                ExportService.ExportAlertsCsv(dialog.FileName, GetEditableAlerts());
            }

            NotificationManager.ShowNotification("CSV Exported", dialog.FileName);
        }

        private void btnPrintPdf_Click(object sender, EventArgs e)
        {
            PrintReportService.PrintReport(
                this,
                GetEditableAlerts(),
                HistoryManager.GetHistory(),
                statusPieChart,
                openPieChart);
        }

        private async void btnTestNtfy_Click(object sender, EventArgs e)
        {
            try
            {
                SaveSettingsFromControls();
                await NtfyService.SendReminderAsync(appSettings, GetEditableAlerts());
                NotificationManager.ShowNotification("ntfy.sh", "Test reminder sent.");
            }
            catch (Exception ex)
            {
                NotificationManager.ShowNotification("ntfy.sh Failed", ex.Message);
            }
        }

        private List<TrainingAlert> GetEditableAlerts()
        {
            dgvAlerts.EndEdit();

            if (dgvAlerts.DataSource is BindingList<TrainingAlert> binding)
            {
                return binding
                    .Where(IsValidAlert)
                    .Select(CloneAlert)
                    .ToList();
            }

            return currentAlerts
                .Where(IsValidAlert)
                .Select(CloneAlert)
                .ToList();
        }

        private static bool IsValidAlert(TrainingAlert alert)
        {
            return alert != null &&
                   !string.IsNullOrWhiteSpace(alert.EmployeeName) &&
                   !string.IsNullOrWhiteSpace(alert.Category) &&
                   !string.IsNullOrWhiteSpace(alert.Status);
        }

        private static TrainingAlert CloneAlert(TrainingAlert alert)
        {
            return new TrainingAlert
            {
                EmployeeName = alert.EmployeeName?.Trim(),
                Category = alert.Category?.Trim(),
                Status = alert.Status?.Trim(),
                Timestamp = string.IsNullOrWhiteSpace(alert.Timestamp)
                    ? DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
                    : alert.Timestamp
            };
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
            if ((cboMinimizeBehavior.SelectedItem?.ToString() ?? "Tray") == "Tray" &&
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
