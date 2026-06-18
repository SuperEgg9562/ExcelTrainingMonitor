using System.Drawing;
using System.Windows.Forms;
using ExcelTrainingMonitor.Controls;

namespace ExcelTrainingMonitor
{
    partial class MainForm
    {
        private System.ComponentModel.IContainer components = null;
        private GlossyButton btnBrowse;
        private GlossyButton btnStart;
        private GlossyButton btnStop;
        private GlossyButton btnNewExcel;
        private GlossyButton btnSaveExcel;
        private GlossyButton btnExportExcel;
        private GlossyButton btnExportCsv;
        private GlossyButton btnPrintPdf;
        private GlossyButton btnTestNtfy;
        private Label lblFile;
        private Label lblInterval;
        private Label lblReminder;
        private Label lblScanHours;
        private Label lblScanMinutes;
        private Label lblTotal;
        private Label lblNotTrained;
        private Label lblTraining;
        private Label lblComplete;
        private NumericUpDown numScanHours;
        private NumericUpDown numScanMinutes;
        private DateTimePicker dtpReminderDate;
        private GlossyCheckBox chkReminderEnabled;
        private GlossyCheckBox chkNtfyEnabled;
        private GlossyComboBox cboTheme;
        private Label lblTheme;
        private Label lblNtfyTopic;
        private Label lblNtfyEmail;
        private TextBox txtNtfyTopic;
        private TextBox txtNtfyEmail;
        private TextBox txtSearch;
        private TabControl tabControl1;
        private TabPage tabMonitor;
        private TabPage tabHistory;
        private TabPage tabCharts;
        private DataGridView dgvAlerts;
        private DataGridView dgvHistory;
        private NotifyIcon notifyIcon1;
        private GlossyCheckBox chkMinimizeTray;
        private NeonProgressBar pbNotTrained;
        private NeonProgressBar pbTraining;
        private NeonProgressBar pbComplete;
        private TableLayoutPanel mainLayout;
        private TableLayoutPanel topLayout;
        private TableLayoutPanel fileSearchLayout;
        private FlowLayoutPanel actionLayout;
        private FlowLayoutPanel dashboardLayout;
        private TableLayoutPanel footerLayout;
        private TableLayoutPanel intervalLayout;
        private TableLayoutPanel reminderLayout;
        private TableLayoutPanel themeLayout;
        private TableLayoutPanel chartsLayout;
        private PieChartPanel statusPieChart;
        private PieChartPanel openPieChart;
        private PictureBox picThemeLogo;
        private PictureBox picAccentBar;
        private TableLayoutPanel titleBarLayout;
        private Label lblWindowTitle;
        private GlossyButton btnWindowMinimize;
        private GlossyButton btnWindowMaximize;
        private GlossyButton btnWindowClose;

        protected override void Dispose(bool disposing)
        {
            if (disposing && components != null)
            {
                components.Dispose();
            }

            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            components = new System.ComponentModel.Container();
            DataGridViewCellStyle dataGridViewCellStyle1 = new DataGridViewCellStyle();
            DataGridViewCellStyle dataGridViewCellStyle2 = new DataGridViewCellStyle();
            btnBrowse = new GlossyButton();
            btnStart = new GlossyButton();
            btnStop = new GlossyButton();
            btnNewExcel = new GlossyButton();
            btnSaveExcel = new GlossyButton();
            btnExportExcel = new GlossyButton();
            btnExportCsv = new GlossyButton();
            btnPrintPdf = new GlossyButton();
            btnTestNtfy = new GlossyButton();
            lblFile = new Label();
            lblInterval = new Label();
            lblReminder = new Label();
            lblScanHours = new Label();
            lblScanMinutes = new Label();
            lblTotal = new Label();
            lblNotTrained = new Label();
            lblTraining = new Label();
            lblComplete = new Label();
            numScanHours = new NumericUpDown();
            numScanMinutes = new NumericUpDown();
            dtpReminderDate = new DateTimePicker();
            chkReminderEnabled = new GlossyCheckBox();
            chkNtfyEnabled = new GlossyCheckBox();
            cboTheme = new GlossyComboBox();
            lblTheme = new Label();
            lblNtfyTopic = new Label();
            lblNtfyEmail = new Label();
            txtNtfyTopic = new TextBox();
            txtNtfyEmail = new TextBox();
            txtSearch = new TextBox();
            tabControl1 = new TabControl();
            tabMonitor = new TabPage();
            dgvAlerts = new DataGridView();
            tabHistory = new TabPage();
            dgvHistory = new DataGridView();
            dataGridViewTextBoxColumn1 = new DataGridViewTextBoxColumn();
            dataGridViewTextBoxColumn2 = new DataGridViewTextBoxColumn();
            dataGridViewTextBoxColumn3 = new DataGridViewTextBoxColumn();
            dataGridViewTextBoxColumn4 = new DataGridViewTextBoxColumn();
            dataGridViewTextBoxColumn5 = new DataGridViewTextBoxColumn();
            tabCharts = new TabPage();
            chartsLayout = new TableLayoutPanel();
            statusPieChart = new PieChartPanel();
            openPieChart = new PieChartPanel();
            notifyIcon1 = new NotifyIcon(components);
            chkMinimizeTray = new GlossyCheckBox();
            pbNotTrained = new NeonProgressBar();
            pbTraining = new NeonProgressBar();
            pbComplete = new NeonProgressBar();
            mainLayout = new TableLayoutPanel();
            titleBarLayout = new TableLayoutPanel();
            lblWindowTitle = new Label();
            btnWindowMinimize = new GlossyButton();
            btnWindowMaximize = new GlossyButton();
            btnWindowClose = new GlossyButton();
            topLayout = new TableLayoutPanel();
            picThemeLogo = new PictureBox();
            intervalLayout = new TableLayoutPanel();
            reminderLayout = new TableLayoutPanel();
            themeLayout = new TableLayoutPanel();
            fileSearchLayout = new TableLayoutPanel();
            actionLayout = new FlowLayoutPanel();
            picAccentBar = new PictureBox();
            dashboardLayout = new FlowLayoutPanel();
            footerLayout = new TableLayoutPanel();
            ((System.ComponentModel.ISupportInitialize)numScanHours).BeginInit();
            ((System.ComponentModel.ISupportInitialize)numScanMinutes).BeginInit();
            tabControl1.SuspendLayout();
            tabMonitor.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)dgvAlerts).BeginInit();
            tabHistory.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)dgvHistory).BeginInit();
            tabCharts.SuspendLayout();
            chartsLayout.SuspendLayout();
            mainLayout.SuspendLayout();
            titleBarLayout.SuspendLayout();
            topLayout.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)picThemeLogo).BeginInit();
            intervalLayout.SuspendLayout();
            reminderLayout.SuspendLayout();
            themeLayout.SuspendLayout();
            fileSearchLayout.SuspendLayout();
            actionLayout.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)picAccentBar).BeginInit();
            dashboardLayout.SuspendLayout();
            footerLayout.SuspendLayout();
            SuspendLayout();
            // 
            // btnBrowse
            // 
            btnBrowse.AutoSize = true;
            btnBrowse.BackColor = Color.FromArgb(3, 45, 13);
            btnBrowse.Dock = DockStyle.Fill;
            btnBrowse.FlatStyle = FlatStyle.Flat;
            btnBrowse.ForeColor = Color.FromArgb(0, 255, 40);
            btnBrowse.Location = new Point(0, 0);
            btnBrowse.Margin = new Padding(0, 0, 19, 0);
            btnBrowse.MinimumSize = new Size(220, 46);
            btnBrowse.Name = "btnBrowse";
            btnBrowse.Size = new Size(239, 120);
            btnBrowse.TabIndex = 0;
            btnBrowse.Text = "Browse Excel";
            btnBrowse.UseVisualStyleBackColor = true;
            btnBrowse.Click += btnBrowse_Click;
            // 
            // btnStart
            // 
            btnStart.BackColor = Color.FromArgb(3, 45, 13);
            btnStart.Dock = DockStyle.Fill;
            btnStart.FlatStyle = FlatStyle.Flat;
            btnStart.ForeColor = Color.FromArgb(0, 255, 40);
            btnStart.Location = new Point(258, 0);
            btnStart.Margin = new Padding(0, 0, 19, 0);
            btnStart.MinimumSize = new Size(157, 46);
            btnStart.Name = "btnStart";
            btnStart.Size = new Size(157, 120);
            btnStart.TabIndex = 1;
            btnStart.Text = "Start";
            btnStart.UseVisualStyleBackColor = true;
            btnStart.Click += btnStart_Click;
            // 
            // btnStop
            // 
            btnStop.BackColor = Color.FromArgb(3, 45, 13);
            btnStop.Dock = DockStyle.Fill;
            btnStop.FlatStyle = FlatStyle.Flat;
            btnStop.ForeColor = Color.FromArgb(0, 255, 40);
            btnStop.Location = new Point(434, 0);
            btnStop.Margin = new Padding(0, 0, 19, 0);
            btnStop.MinimumSize = new Size(157, 46);
            btnStop.Name = "btnStop";
            btnStop.Size = new Size(157, 120);
            btnStop.TabIndex = 2;
            btnStop.Text = "Stop";
            btnStop.UseVisualStyleBackColor = true;
            btnStop.Click += btnStop_Click;
            // 
            // btnNewExcel
            // 
            btnNewExcel.AutoSize = true;
            btnNewExcel.BackColor = Color.FromArgb(3, 45, 13);
            btnNewExcel.FlatStyle = FlatStyle.Flat;
            btnNewExcel.ForeColor = Color.FromArgb(0, 255, 40);
            btnNewExcel.Location = new Point(0, 0);
            btnNewExcel.Margin = new Padding(0, 0, 13, 7);
            btnNewExcel.MinimumSize = new Size(163, 41);
            btnNewExcel.Name = "btnNewExcel";
            btnNewExcel.Size = new Size(187, 41);
            btnNewExcel.TabIndex = 0;
            btnNewExcel.Text = "New Excel";
            btnNewExcel.UseVisualStyleBackColor = false;
            btnNewExcel.Click += btnNewExcel_Click;
            // 
            // btnSaveExcel
            // 
            btnSaveExcel.AutoSize = true;
            btnSaveExcel.BackColor = Color.FromArgb(3, 45, 13);
            btnSaveExcel.FlatStyle = FlatStyle.Flat;
            btnSaveExcel.ForeColor = Color.FromArgb(0, 255, 40);
            btnSaveExcel.Location = new Point(200, 0);
            btnSaveExcel.Margin = new Padding(0, 0, 13, 7);
            btnSaveExcel.MinimumSize = new Size(176, 41);
            btnSaveExcel.Name = "btnSaveExcel";
            btnSaveExcel.Size = new Size(204, 41);
            btnSaveExcel.TabIndex = 1;
            btnSaveExcel.Text = "Save Edits";
            btnSaveExcel.UseVisualStyleBackColor = false;
            btnSaveExcel.Click += btnSaveExcel_Click;
            // 
            // btnExportExcel
            // 
            btnExportExcel.AutoSize = true;
            btnExportExcel.BackColor = Color.FromArgb(3, 45, 13);
            btnExportExcel.FlatStyle = FlatStyle.Flat;
            btnExportExcel.ForeColor = Color.FromArgb(0, 255, 40);
            btnExportExcel.Location = new Point(417, 0);
            btnExportExcel.Margin = new Padding(0, 0, 13, 7);
            btnExportExcel.MinimumSize = new Size(185, 41);
            btnExportExcel.Name = "btnExportExcel";
            btnExportExcel.Size = new Size(239, 41);
            btnExportExcel.TabIndex = 2;
            btnExportExcel.Text = "Export Excel";
            btnExportExcel.UseVisualStyleBackColor = false;
            btnExportExcel.Click += btnExportExcel_Click;
            // 
            // btnExportCsv
            // 
            btnExportCsv.AutoSize = true;
            btnExportCsv.BackColor = Color.FromArgb(3, 45, 13);
            btnExportCsv.FlatStyle = FlatStyle.Flat;
            btnExportCsv.ForeColor = Color.FromArgb(0, 255, 40);
            btnExportCsv.Location = new Point(669, 0);
            btnExportCsv.Margin = new Padding(0, 0, 13, 7);
            btnExportCsv.MinimumSize = new Size(163, 41);
            btnExportCsv.Name = "btnExportCsv";
            btnExportCsv.Size = new Size(204, 41);
            btnExportCsv.TabIndex = 3;
            btnExportCsv.Text = "Export CSV";
            btnExportCsv.UseVisualStyleBackColor = false;
            btnExportCsv.Click += btnExportCsv_Click;
            // 
            // btnPrintPdf
            // 
            btnPrintPdf.AutoSize = true;
            btnPrintPdf.BackColor = Color.FromArgb(3, 45, 13);
            btnPrintPdf.FlatStyle = FlatStyle.Flat;
            btnPrintPdf.ForeColor = Color.FromArgb(0, 255, 40);
            btnPrintPdf.Location = new Point(886, 0);
            btnPrintPdf.Margin = new Padding(0, 0, 19, 7);
            btnPrintPdf.MinimumSize = new Size(163, 41);
            btnPrintPdf.Name = "btnPrintPdf";
            btnPrintPdf.Size = new Size(187, 41);
            btnPrintPdf.TabIndex = 4;
            btnPrintPdf.Text = "Print PDF";
            btnPrintPdf.UseVisualStyleBackColor = false;
            btnPrintPdf.Click += btnPrintPdf_Click;
            // 
            // btnTestNtfy
            // 
            btnTestNtfy.AutoSize = true;
            btnTestNtfy.BackColor = Color.FromArgb(3, 45, 13);
            btnTestNtfy.FlatStyle = FlatStyle.Flat;
            btnTestNtfy.ForeColor = Color.FromArgb(0, 255, 40);
            btnTestNtfy.Location = new Point(312, 48);
            btnTestNtfy.Margin = new Padding(0, 0, 0, 7);
            btnTestNtfy.MinimumSize = new Size(154, 41);
            btnTestNtfy.Name = "btnTestNtfy";
            btnTestNtfy.Size = new Size(187, 41);
            btnTestNtfy.TabIndex = 10;
            btnTestNtfy.Text = "Test ntfy";
            btnTestNtfy.UseVisualStyleBackColor = false;
            btnTestNtfy.Click += btnTestNtfy_Click;
            // 
            // lblFile
            // 
            lblFile.AutoEllipsis = true;
            lblFile.Dock = DockStyle.Fill;
            lblFile.Location = new Point(0, 7);
            lblFile.Margin = new Padding(0, 7, 25, 0);
            lblFile.Name = "lblFile";
            lblFile.Size = new Size(1439, 113);
            lblFile.TabIndex = 3;
            lblFile.Text = "No Excel file selected";
            lblFile.TextAlign = ContentAlignment.MiddleLeft;
            // 
            // lblInterval
            // 
            lblInterval.AutoSize = true;
            intervalLayout.SetColumnSpan(lblInterval, 2);
            lblInterval.Dock = DockStyle.Fill;
            lblInterval.Location = new Point(5, 0);
            lblInterval.Margin = new Padding(5, 0, 5, 0);
            lblInterval.Name = "lblInterval";
            lblInterval.Size = new Size(289, 18);
            lblInterval.TabIndex = 0;
            lblInterval.Text = "Scan Interval";
            lblInterval.TextAlign = ContentAlignment.BottomLeft;
            // 
            // lblReminder
            // 
            lblReminder.AutoSize = true;
            lblReminder.Dock = DockStyle.Fill;
            lblReminder.Location = new Point(5, 0);
            lblReminder.Margin = new Padding(5, 0, 5, 0);
            lblReminder.Name = "lblReminder";
            lblReminder.Size = new Size(351, 18);
            lblReminder.TabIndex = 0;
            lblReminder.Text = "Reminder Date";
            lblReminder.TextAlign = ContentAlignment.BottomLeft;
            // 
            // lblScanHours
            // 
            lblScanHours.AutoSize = true;
            lblScanHours.Dock = DockStyle.Fill;
            lblScanHours.Location = new Point(5, 18);
            lblScanHours.Margin = new Padding(5, 0, 5, 0);
            lblScanHours.Name = "lblScanHours";
            lblScanHours.Size = new Size(139, 18);
            lblScanHours.TabIndex = 1;
            lblScanHours.Text = "Hours";
            // 
            // lblScanMinutes
            // 
            lblScanMinutes.AutoSize = true;
            lblScanMinutes.Dock = DockStyle.Fill;
            lblScanMinutes.Location = new Point(154, 18);
            lblScanMinutes.Margin = new Padding(5, 0, 5, 0);
            lblScanMinutes.Name = "lblScanMinutes";
            lblScanMinutes.Size = new Size(140, 18);
            lblScanMinutes.TabIndex = 2;
            lblScanMinutes.Text = "Minutes";
            // 
            // lblTotal
            // 
            lblTotal.AutoSize = true;
            lblTotal.Font = new Font("Segoe UI", 10F, FontStyle.Bold);
            lblTotal.Location = new Point(0, 0);
            lblTotal.Margin = new Padding(0, 0, 38, 0);
            lblTotal.Name = "lblTotal";
            lblTotal.Size = new Size(113, 20);
            lblTotal.TabIndex = 0;
            lblTotal.Text = "Total: 0";
            // 
            // lblNotTrained
            // 
            lblNotTrained.AutoSize = true;
            lblNotTrained.ForeColor = Color.FromArgb(255, 70, 70);
            lblNotTrained.Location = new Point(151, 0);
            lblNotTrained.Margin = new Padding(0, 0, 38, 0);
            lblNotTrained.Name = "lblNotTrained";
            lblNotTrained.Size = new Size(162, 18);
            lblNotTrained.TabIndex = 1;
            lblNotTrained.Text = "Not Trained: 0";
            // 
            // lblTraining
            // 
            lblTraining.AutoSize = true;
            lblTraining.ForeColor = Color.FromArgb(255, 210, 28);
            lblTraining.Location = new Point(351, 0);
            lblTraining.Margin = new Padding(0, 0, 38, 0);
            lblTraining.Name = "lblTraining";
            lblTraining.Size = new Size(162, 18);
            lblTraining.TabIndex = 2;
            lblTraining.Text = "In Training: 0";
            // 
            // lblComplete
            // 
            lblComplete.AutoSize = true;
            lblComplete.ForeColor = Color.FromArgb(0, 255, 40);
            lblComplete.Location = new Point(551, 0);
            lblComplete.Margin = new Padding(0);
            lblComplete.Name = "lblComplete";
            lblComplete.Size = new Size(129, 18);
            lblComplete.TabIndex = 3;
            lblComplete.Text = "Complete: 0";
            // 
            // numScanHours
            // 
            numScanHours.Dock = DockStyle.Fill;
            numScanHours.Location = new Point(5, 40);
            numScanHours.Margin = new Padding(5, 4, 5, 4);
            numScanHours.Maximum = new decimal(new int[] { 24, 0, 0, 0 });
            numScanHours.Name = "numScanHours";
            numScanHours.Size = new Size(139, 31);
            numScanHours.TabIndex = 5;
            // 
            // numScanMinutes
            // 
            numScanMinutes.Dock = DockStyle.Fill;
            numScanMinutes.Location = new Point(154, 40);
            numScanMinutes.Margin = new Padding(5, 4, 5, 4);
            numScanMinutes.Maximum = new decimal(new int[] { 59, 0, 0, 0 });
            numScanMinutes.Name = "numScanMinutes";
            numScanMinutes.Size = new Size(140, 31);
            numScanMinutes.TabIndex = 6;
            numScanMinutes.Value = new decimal(new int[] { 5, 0, 0, 0 });
            // 
            // dtpReminderDate
            // 
            dtpReminderDate.CustomFormat = "yyyy-MM-dd HH:mm";
            dtpReminderDate.Dock = DockStyle.Fill;
            dtpReminderDate.Format = DateTimePickerFormat.Custom;
            dtpReminderDate.Location = new Point(5, 22);
            dtpReminderDate.Margin = new Padding(5, 4, 5, 4);
            dtpReminderDate.Name = "dtpReminderDate";
            dtpReminderDate.Size = new Size(351, 31);
            dtpReminderDate.TabIndex = 1;
            // 
            // chkReminderEnabled
            // 
            chkReminderEnabled.AutoSize = true;
            chkReminderEnabled.BackColor = Color.Transparent;
            chkReminderEnabled.Dock = DockStyle.Fill;
            chkReminderEnabled.FlatStyle = FlatStyle.Flat;
            chkReminderEnabled.ForeColor = Color.FromArgb(0, 255, 40);
            chkReminderEnabled.Location = new Point(5, 61);
            chkReminderEnabled.Margin = new Padding(5, 4, 5, 4);
            chkReminderEnabled.Name = "chkReminderEnabled";
            chkReminderEnabled.Size = new Size(351, 55);
            chkReminderEnabled.TabIndex = 2;
            chkReminderEnabled.Text = "Enable reminder";
            chkReminderEnabled.UseVisualStyleBackColor = false;
            // 
            // chkNtfyEnabled
            // 
            chkNtfyEnabled.AutoSize = true;
            chkNtfyEnabled.BackColor = Color.Transparent;
            chkNtfyEnabled.FlatStyle = FlatStyle.Flat;
            chkNtfyEnabled.ForeColor = Color.FromArgb(0, 255, 40);
            chkNtfyEnabled.Location = new Point(1092, 7);
            chkNtfyEnabled.Margin = new Padding(0, 7, 16, 7);
            chkNtfyEnabled.Name = "chkNtfyEnabled";
            chkNtfyEnabled.Size = new Size(106, 22);
            chkNtfyEnabled.TabIndex = 5;
            chkNtfyEnabled.Text = "ntfy.sh";
            chkNtfyEnabled.UseVisualStyleBackColor = false;
            // 
            // cboTheme
            // 
            cboTheme.BackColor = Color.FromArgb(2, 34, 10);
            cboTheme.Dock = DockStyle.Fill;
            cboTheme.DrawMode = DrawMode.OwnerDrawFixed;
            cboTheme.DropDownStyle = ComboBoxStyle.DropDownList;
            cboTheme.FlatStyle = FlatStyle.Flat;
            cboTheme.ForeColor = Color.FromArgb(0, 255, 40);
            cboTheme.ItemHeight = 24;
            cboTheme.Items.AddRange(new object[] { "Dark" });
            cboTheme.Location = new Point(5, 22);
            cboTheme.Margin = new Padding(5, 4, 5, 4);
            cboTheme.Name = "cboTheme";
            cboTheme.Size = new Size(210, 30);
            cboTheme.TabIndex = 1;
            // 
            // lblTheme
            // 
            lblTheme.AutoSize = true;
            lblTheme.Dock = DockStyle.Fill;
            lblTheme.Location = new Point(5, 0);
            lblTheme.Margin = new Padding(5, 0, 5, 0);
            lblTheme.Name = "lblTheme";
            lblTheme.Size = new Size(210, 18);
            lblTheme.TabIndex = 0;
            lblTheme.Text = "Theme";
            lblTheme.TextAlign = ContentAlignment.BottomLeft;
            // 
            // lblNtfyTopic
            // 
            lblNtfyTopic.AutoSize = true;
            lblNtfyTopic.Location = new Point(1214, 10);
            lblNtfyTopic.Margin = new Padding(0, 10, 6, 0);
            lblNtfyTopic.Name = "lblNtfyTopic";
            lblNtfyTopic.Size = new Size(63, 18);
            lblNtfyTopic.TabIndex = 6;
            lblNtfyTopic.Text = "Topic";
            // 
            // lblNtfyEmail
            // 
            lblNtfyEmail.AutoSize = true;
            lblNtfyEmail.Location = new Point(1548, 10);
            lblNtfyEmail.Margin = new Padding(0, 10, 6, 0);
            lblNtfyEmail.Name = "lblNtfyEmail";
            lblNtfyEmail.Size = new Size(63, 18);
            lblNtfyEmail.TabIndex = 8;
            lblNtfyEmail.Text = "Email";
            // 
            // txtNtfyTopic
            // 
            txtNtfyTopic.Location = new Point(1283, 5);
            txtNtfyTopic.Margin = new Padding(0, 5, 16, 7);
            txtNtfyTopic.Name = "txtNtfyTopic";
            txtNtfyTopic.Size = new Size(249, 31);
            txtNtfyTopic.TabIndex = 7;
            // 
            // txtNtfyEmail
            // 
            txtNtfyEmail.Location = new Point(0, 53);
            txtNtfyEmail.Margin = new Padding(0, 5, 16, 7);
            txtNtfyEmail.Name = "txtNtfyEmail";
            txtNtfyEmail.Size = new Size(296, 31);
            txtNtfyEmail.TabIndex = 9;
            // 
            // txtSearch
            // 
            txtSearch.Dock = DockStyle.Fill;
            txtSearch.Location = new Point(1464, 7);
            txtSearch.Margin = new Padding(0, 7, 0, 0);
            txtSearch.Name = "txtSearch";
            txtSearch.PlaceholderText = "Search";
            txtSearch.Size = new Size(440, 31);
            txtSearch.TabIndex = 8;
            txtSearch.TextChanged += TxtSearch_TextChanged;
            // 
            // tabControl1
            // 
            tabControl1.Controls.Add(tabMonitor);
            tabControl1.Controls.Add(tabHistory);
            tabControl1.Controls.Add(tabCharts);
            tabControl1.Dock = DockStyle.Fill;
            tabControl1.Location = new Point(21, 464);
            tabControl1.Margin = new Padding(19, 0, 19, 14);
            tabControl1.Name = "tabControl1";
            tabControl1.SelectedIndex = 0;
            tabControl1.Size = new Size(1904, 372);
            tabControl1.TabIndex = 6;
            // 
            // tabMonitor
            // 
            tabMonitor.Controls.Add(dgvAlerts);
            tabMonitor.Location = new Point(4, 28);
            tabMonitor.Margin = new Padding(5, 4, 5, 4);
            tabMonitor.Name = "tabMonitor";
            tabMonitor.Padding = new Padding(5, 4, 5, 4);
            tabMonitor.Size = new Size(1896, 340);
            tabMonitor.TabIndex = 0;
            tabMonitor.Text = "Monitor";
            // 
            // dgvAlerts
            // 
            dgvAlerts.AllowUserToAddRows = false;
            dgvAlerts.AllowUserToDeleteRows = false;
            dgvAlerts.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dgvAlerts.BackgroundColor = Color.FromArgb(8, 10, 9);
            dgvAlerts.BorderStyle = BorderStyle.None;
            dataGridViewCellStyle1.Alignment = DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = Color.FromArgb(3, 45, 13);
            dataGridViewCellStyle1.Font = new Font("Segoe UI", 10F);
            dataGridViewCellStyle1.ForeColor = Color.FromArgb(0, 255, 40);
            dataGridViewCellStyle1.SelectionBackColor = Color.FromArgb(0, 220, 35);
            dataGridViewCellStyle1.SelectionForeColor = Color.Black;
            dataGridViewCellStyle1.WrapMode = DataGridViewTriState.True;
            dgvAlerts.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
            dgvAlerts.ColumnHeadersHeight = 34;
            dataGridViewCellStyle2.Alignment = DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.BackColor = Color.FromArgb(8, 10, 9);
            dataGridViewCellStyle2.Font = new Font("Segoe UI", 10F);
            dataGridViewCellStyle2.ForeColor = Color.FromArgb(0, 255, 40);
            dataGridViewCellStyle2.SelectionBackColor = Color.FromArgb(0, 220, 35);
            dataGridViewCellStyle2.SelectionForeColor = Color.Black;
            dataGridViewCellStyle2.WrapMode = DataGridViewTriState.False;
            dgvAlerts.DefaultCellStyle = dataGridViewCellStyle2;
            dgvAlerts.Dock = DockStyle.Fill;
            dgvAlerts.EnableHeadersVisualStyles = false;
            dgvAlerts.Font = new Font("Segoe UI", 10F);
            dgvAlerts.Location = new Point(5, 4);
            dgvAlerts.Margin = new Padding(0);
            dgvAlerts.Name = "dgvAlerts";
            dgvAlerts.ReadOnly = true;
            dgvAlerts.RowHeadersVisible = false;
            dgvAlerts.RowHeadersWidth = 62;
            dgvAlerts.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dgvAlerts.Size = new Size(1886, 332);
            dgvAlerts.TabIndex = 9;
            // 
            // tabHistory
            // 
            tabHistory.Controls.Add(dgvHistory);
            tabHistory.Location = new Point(4, 28);
            tabHistory.Margin = new Padding(5, 4, 5, 4);
            tabHistory.Name = "tabHistory";
            tabHistory.Padding = new Padding(5, 4, 5, 4);
            tabHistory.Size = new Size(265, 0);
            tabHistory.TabIndex = 1;
            tabHistory.Text = "History";
            // 
            // dgvHistory
            // 
            dgvHistory.AllowUserToAddRows = false;
            dgvHistory.AllowUserToDeleteRows = false;
            dgvHistory.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dgvHistory.ColumnHeadersHeight = 34;
            dgvHistory.Columns.AddRange(new DataGridViewColumn[] { dataGridViewTextBoxColumn1, dataGridViewTextBoxColumn2, dataGridViewTextBoxColumn3, dataGridViewTextBoxColumn4, dataGridViewTextBoxColumn5 });
            dgvHistory.Dock = DockStyle.Fill;
            dgvHistory.Location = new Point(5, 4);
            dgvHistory.Margin = new Padding(0);
            dgvHistory.Name = "dgvHistory";
            dgvHistory.ReadOnly = true;
            dgvHistory.RowHeadersVisible = false;
            dgvHistory.RowHeadersWidth = 62;
            dgvHistory.Size = new Size(255, 0);
            dgvHistory.TabIndex = 0;
            // 
            // dataGridViewTextBoxColumn1
            // 
            dataGridViewTextBoxColumn1.MinimumWidth = 8;
            dataGridViewTextBoxColumn1.Name = "dataGridViewTextBoxColumn1";
            dataGridViewTextBoxColumn1.ReadOnly = true;
            // 
            // dataGridViewTextBoxColumn2
            // 
            dataGridViewTextBoxColumn2.MinimumWidth = 8;
            dataGridViewTextBoxColumn2.Name = "dataGridViewTextBoxColumn2";
            dataGridViewTextBoxColumn2.ReadOnly = true;
            // 
            // dataGridViewTextBoxColumn3
            // 
            dataGridViewTextBoxColumn3.MinimumWidth = 8;
            dataGridViewTextBoxColumn3.Name = "dataGridViewTextBoxColumn3";
            dataGridViewTextBoxColumn3.ReadOnly = true;
            // 
            // dataGridViewTextBoxColumn4
            // 
            dataGridViewTextBoxColumn4.MinimumWidth = 8;
            dataGridViewTextBoxColumn4.Name = "dataGridViewTextBoxColumn4";
            dataGridViewTextBoxColumn4.ReadOnly = true;
            // 
            // dataGridViewTextBoxColumn5
            // 
            dataGridViewTextBoxColumn5.MinimumWidth = 8;
            dataGridViewTextBoxColumn5.Name = "dataGridViewTextBoxColumn5";
            dataGridViewTextBoxColumn5.ReadOnly = true;
            // 
            // tabCharts
            // 
            tabCharts.Controls.Add(chartsLayout);
            tabCharts.Location = new Point(4, 28);
            tabCharts.Margin = new Padding(5, 4, 5, 4);
            tabCharts.Name = "tabCharts";
            tabCharts.Padding = new Padding(5, 4, 5, 4);
            tabCharts.Size = new Size(265, 0);
            tabCharts.TabIndex = 2;
            tabCharts.Text = "Graphs";
            // 
            // chartsLayout
            // 
            chartsLayout.ColumnCount = 2;
            chartsLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
            chartsLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
            chartsLayout.Controls.Add(statusPieChart, 0, 0);
            chartsLayout.Controls.Add(openPieChart, 1, 0);
            chartsLayout.Dock = DockStyle.Fill;
            chartsLayout.Location = new Point(5, 4);
            chartsLayout.Margin = new Padding(0);
            chartsLayout.Name = "chartsLayout";
            chartsLayout.RowCount = 1;
            chartsLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            chartsLayout.Size = new Size(255, 0);
            chartsLayout.TabIndex = 0;
            // 
            // statusPieChart
            // 
            statusPieChart.ChartTitle = "All Training Statuses";
            statusPieChart.Dock = DockStyle.Fill;
            statusPieChart.Font = new Font("Segoe UI", 10F);
            statusPieChart.Location = new Point(0, 0);
            statusPieChart.Margin = new Padding(0, 0, 13, 0);
            statusPieChart.MinimumSize = new Size(346, 264);
            statusPieChart.Name = "statusPieChart";
            statusPieChart.Size = new Size(346, 264);
            statusPieChart.TabIndex = 0;
            // 
            // openPieChart
            // 
            openPieChart.ChartTitle = "Open Training Items";
            openPieChart.Dock = DockStyle.Fill;
            openPieChart.Font = new Font("Segoe UI", 10F);
            openPieChart.Location = new Point(140, 0);
            openPieChart.Margin = new Padding(13, 0, 0, 0);
            openPieChart.MinimumSize = new Size(346, 264);
            openPieChart.Name = "openPieChart";
            openPieChart.Size = new Size(346, 264);
            openPieChart.TabIndex = 1;
            // 
            // notifyIcon1
            // 
            notifyIcon1.Text = "Excel Training Monitor";
            notifyIcon1.Visible = true;
            notifyIcon1.DoubleClick += notifyIcon1_DoubleClick;
            // 
            // chkMinimizeTray
            // 
            chkMinimizeTray.AutoSize = true;
            chkMinimizeTray.BackColor = Color.Transparent;
            chkMinimizeTray.Checked = true;
            chkMinimizeTray.CheckState = CheckState.Checked;
            chkMinimizeTray.Dock = DockStyle.Right;
            chkMinimizeTray.FlatStyle = FlatStyle.Flat;
            chkMinimizeTray.ForeColor = Color.FromArgb(0, 255, 40);
            chkMinimizeTray.Location = new Point(1699, 0);
            chkMinimizeTray.Margin = new Padding(0);
            chkMinimizeTray.Name = "chkMinimizeTray";
            chkMinimizeTray.Size = new Size(205, 120);
            chkMinimizeTray.TabIndex = 10;
            chkMinimizeTray.Text = "Minimize to tray";
            chkMinimizeTray.UseVisualStyleBackColor = false;
            // 
            // pbNotTrained
            // 
            pbNotTrained.Dock = DockStyle.Fill;
            pbNotTrained.Location = new Point(0, 0);
            pbNotTrained.Margin = new Padding(0, 0, 19, 0);
            pbNotTrained.Maximum = 100;
            pbNotTrained.MinimumSize = new Size(126, 22);
            pbNotTrained.Name = "pbNotTrained";
            pbNotTrained.Size = new Size(405, 120);
            pbNotTrained.TabIndex = 0;
            pbNotTrained.Value = 0;
            // 
            // pbTraining
            // 
            pbTraining.Dock = DockStyle.Fill;
            pbTraining.Location = new Point(424, 0);
            pbTraining.Margin = new Padding(0, 0, 19, 0);
            pbTraining.Maximum = 100;
            pbTraining.MinimumSize = new Size(126, 22);
            pbTraining.Name = "pbTraining";
            pbTraining.Size = new Size(405, 120);
            pbTraining.TabIndex = 1;
            pbTraining.Value = 0;
            // 
            // pbComplete
            // 
            pbComplete.Dock = DockStyle.Fill;
            pbComplete.Location = new Point(848, 0);
            pbComplete.Margin = new Padding(0, 0, 38, 0);
            pbComplete.Maximum = 100;
            pbComplete.MinimumSize = new Size(126, 22);
            pbComplete.Name = "pbComplete";
            pbComplete.Size = new Size(386, 120);
            pbComplete.TabIndex = 2;
            pbComplete.Value = 0;
            // 
            // mainLayout
            // 
            mainLayout.ColumnCount = 1;
            mainLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            mainLayout.Controls.Add(titleBarLayout, 0, 0);
            mainLayout.Controls.Add(topLayout, 0, 1);
            mainLayout.Controls.Add(fileSearchLayout, 0, 2);
            mainLayout.Controls.Add(actionLayout, 0, 3);
            mainLayout.Controls.Add(picAccentBar, 0, 4);
            mainLayout.Controls.Add(dashboardLayout, 0, 5);
            mainLayout.Controls.Add(tabControl1, 0, 6);
            mainLayout.Controls.Add(footerLayout, 0, 7);
            mainLayout.Dock = DockStyle.Fill;
            mainLayout.Location = new Point(0, 0);
            mainLayout.Margin = new Padding(0);
            mainLayout.Name = "mainLayout";
            mainLayout.Padding = new Padding(2, 1, 2, 14);
            mainLayout.RowCount = 8;
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 46F));
            mainLayout.RowStyles.Add(new RowStyle());
            mainLayout.RowStyles.Add(new RowStyle());
            mainLayout.RowStyles.Add(new RowStyle());
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 14F));
            mainLayout.RowStyles.Add(new RowStyle());
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            mainLayout.RowStyles.Add(new RowStyle());
            mainLayout.Size = new Size(1946, 984);
            mainLayout.TabIndex = 0;
            // 
            // titleBarLayout
            // 
            titleBarLayout.ColumnCount = 4;
            titleBarLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            titleBarLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 69F));
            titleBarLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 69F));
            titleBarLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 69F));
            titleBarLayout.Controls.Add(lblWindowTitle, 0, 0);
            titleBarLayout.Controls.Add(btnWindowMinimize, 1, 0);
            titleBarLayout.Controls.Add(btnWindowMaximize, 2, 0);
            titleBarLayout.Controls.Add(btnWindowClose, 3, 0);
            titleBarLayout.Dock = DockStyle.Fill;
            titleBarLayout.Location = new Point(2, 1);
            titleBarLayout.Margin = new Padding(0);
            titleBarLayout.Name = "titleBarLayout";
            titleBarLayout.RowCount = 1;
            titleBarLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            titleBarLayout.Size = new Size(1942, 46);
            titleBarLayout.TabIndex = 0;
            titleBarLayout.MouseDown += TitleBar_MouseDown;
            // 
            // lblWindowTitle
            // 
            lblWindowTitle.AutoSize = true;
            lblWindowTitle.Dock = DockStyle.Fill;
            lblWindowTitle.Font = new Font("Consolas", 11F, FontStyle.Bold);
            lblWindowTitle.Location = new Point(16, 0);
            lblWindowTitle.Margin = new Padding(16, 0, 0, 0);
            lblWindowTitle.Name = "lblWindowTitle";
            lblWindowTitle.Size = new Size(1719, 46);
            lblWindowTitle.TabIndex = 0;
            lblWindowTitle.Text = "Excel Training Monitor";
            lblWindowTitle.TextAlign = ContentAlignment.MiddleLeft;
            lblWindowTitle.MouseDown += TitleBar_MouseDown;
            // 
            // btnWindowMinimize
            // 
            btnWindowMinimize.BackColor = Color.FromArgb(3, 45, 13);
            btnWindowMinimize.Dock = DockStyle.Fill;
            btnWindowMinimize.FlatStyle = FlatStyle.Flat;
            btnWindowMinimize.ForeColor = Color.FromArgb(0, 255, 40);
            btnWindowMinimize.Location = new Point(1741, 4);
            btnWindowMinimize.Margin = new Padding(6, 4, 0, 4);
            btnWindowMinimize.Name = "btnWindowMinimize";
            btnWindowMinimize.Size = new Size(63, 38);
            btnWindowMinimize.TabIndex = 1;
            btnWindowMinimize.Text = "_";
            btnWindowMinimize.UseVisualStyleBackColor = false;
            btnWindowMinimize.Click += btnWindowMinimize_Click;
            // 
            // btnWindowMaximize
            // 
            btnWindowMaximize.BackColor = Color.FromArgb(3, 45, 13);
            btnWindowMaximize.Dock = DockStyle.Fill;
            btnWindowMaximize.FlatStyle = FlatStyle.Flat;
            btnWindowMaximize.ForeColor = Color.FromArgb(0, 255, 40);
            btnWindowMaximize.Location = new Point(1810, 4);
            btnWindowMaximize.Margin = new Padding(6, 4, 0, 4);
            btnWindowMaximize.Name = "btnWindowMaximize";
            btnWindowMaximize.Size = new Size(63, 38);
            btnWindowMaximize.TabIndex = 2;
            btnWindowMaximize.Text = "□";
            btnWindowMaximize.UseVisualStyleBackColor = false;
            btnWindowMaximize.Click += btnWindowMaximize_Click;
            // 
            // btnWindowClose
            // 
            btnWindowClose.BackColor = Color.FromArgb(3, 45, 13);
            btnWindowClose.Dock = DockStyle.Fill;
            btnWindowClose.FlatStyle = FlatStyle.Flat;
            btnWindowClose.ForeColor = Color.FromArgb(0, 255, 40);
            btnWindowClose.Location = new Point(1879, 4);
            btnWindowClose.Margin = new Padding(6, 4, 6, 4);
            btnWindowClose.Name = "btnWindowClose";
            btnWindowClose.Size = new Size(57, 38);
            btnWindowClose.TabIndex = 3;
            btnWindowClose.Text = "X";
            btnWindowClose.UseVisualStyleBackColor = false;
            btnWindowClose.Click += btnWindowClose_Click;
            // 
            // topLayout
            // 
            topLayout.ColumnCount = 7;
            topLayout.ColumnStyles.Add(new ColumnStyle());
            topLayout.ColumnStyles.Add(new ColumnStyle());
            topLayout.ColumnStyles.Add(new ColumnStyle());
            topLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            topLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 299F));
            topLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 361F));
            topLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 220F));
            topLayout.Controls.Add(btnBrowse, 0, 0);
            topLayout.Controls.Add(btnStart, 1, 0);
            topLayout.Controls.Add(btnStop, 2, 0);
            topLayout.Controls.Add(picThemeLogo, 3, 0);
            topLayout.Controls.Add(intervalLayout, 4, 0);
            topLayout.Controls.Add(reminderLayout, 5, 0);
            topLayout.Controls.Add(themeLayout, 6, 0);
            topLayout.Dock = DockStyle.Fill;
            topLayout.Location = new Point(21, 59);
            topLayout.Margin = new Padding(19, 12, 19, 0);
            topLayout.Name = "topLayout";
            topLayout.RowCount = 1;
            topLayout.RowStyles.Add(new RowStyle());
            topLayout.Size = new Size(1904, 120);
            topLayout.TabIndex = 1;
            // 
            // picThemeLogo
            // 
            picThemeLogo.Dock = DockStyle.Right;
            picThemeLogo.Location = new Point(997, 0);
            picThemeLogo.Margin = new Padding(0, 0, 25, 0);
            picThemeLogo.Name = "picThemeLogo";
            picThemeLogo.Size = new Size(2, 120);
            picThemeLogo.SizeMode = PictureBoxSizeMode.Zoom;
            picThemeLogo.TabIndex = 3;
            picThemeLogo.TabStop = false;
            // 
            // intervalLayout
            // 
            intervalLayout.ColumnCount = 2;
            intervalLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
            intervalLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
            intervalLayout.Controls.Add(lblInterval, 0, 0);
            intervalLayout.Controls.Add(lblScanHours, 0, 1);
            intervalLayout.Controls.Add(lblScanMinutes, 1, 1);
            intervalLayout.Controls.Add(numScanHours, 0, 2);
            intervalLayout.Controls.Add(numScanMinutes, 1, 2);
            intervalLayout.Dock = DockStyle.Fill;
            intervalLayout.Location = new Point(1024, 0);
            intervalLayout.Margin = new Padding(0);
            intervalLayout.Name = "intervalLayout";
            intervalLayout.RowCount = 3;
            intervalLayout.RowStyles.Add(new RowStyle());
            intervalLayout.RowStyles.Add(new RowStyle());
            intervalLayout.RowStyles.Add(new RowStyle());
            intervalLayout.Size = new Size(299, 120);
            intervalLayout.TabIndex = 4;
            // 
            // reminderLayout
            // 
            reminderLayout.ColumnCount = 1;
            reminderLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            reminderLayout.Controls.Add(lblReminder, 0, 0);
            reminderLayout.Controls.Add(dtpReminderDate, 0, 1);
            reminderLayout.Controls.Add(chkReminderEnabled, 0, 2);
            reminderLayout.Dock = DockStyle.Fill;
            reminderLayout.Location = new Point(1323, 0);
            reminderLayout.Margin = new Padding(0);
            reminderLayout.Name = "reminderLayout";
            reminderLayout.RowCount = 3;
            reminderLayout.RowStyles.Add(new RowStyle());
            reminderLayout.RowStyles.Add(new RowStyle());
            reminderLayout.RowStyles.Add(new RowStyle());
            reminderLayout.Size = new Size(361, 120);
            reminderLayout.TabIndex = 5;
            // 
            // themeLayout
            // 
            themeLayout.ColumnCount = 1;
            themeLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            themeLayout.Controls.Add(lblTheme, 0, 0);
            themeLayout.Controls.Add(cboTheme, 0, 1);
            themeLayout.Dock = DockStyle.Fill;
            themeLayout.Location = new Point(1684, 0);
            themeLayout.Margin = new Padding(0);
            themeLayout.Name = "themeLayout";
            themeLayout.RowCount = 2;
            themeLayout.RowStyles.Add(new RowStyle());
            themeLayout.RowStyles.Add(new RowStyle());
            themeLayout.Size = new Size(220, 120);
            themeLayout.TabIndex = 6;
            // 
            // fileSearchLayout
            // 
            fileSearchLayout.ColumnCount = 2;
            fileSearchLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            fileSearchLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 440F));
            fileSearchLayout.Controls.Add(lblFile, 0, 0);
            fileSearchLayout.Controls.Add(txtSearch, 1, 0);
            fileSearchLayout.Dock = DockStyle.Fill;
            fileSearchLayout.Location = new Point(21, 189);
            fileSearchLayout.Margin = new Padding(19, 10, 19, 0);
            fileSearchLayout.Name = "fileSearchLayout";
            fileSearchLayout.RowCount = 1;
            fileSearchLayout.RowStyles.Add(new RowStyle());
            fileSearchLayout.Size = new Size(1904, 120);
            fileSearchLayout.TabIndex = 2;
            // 
            // actionLayout
            // 
            actionLayout.AutoSize = true;
            actionLayout.Controls.Add(btnNewExcel);
            actionLayout.Controls.Add(btnSaveExcel);
            actionLayout.Controls.Add(btnExportExcel);
            actionLayout.Controls.Add(btnExportCsv);
            actionLayout.Controls.Add(btnPrintPdf);
            actionLayout.Controls.Add(chkNtfyEnabled);
            actionLayout.Controls.Add(lblNtfyTopic);
            actionLayout.Controls.Add(txtNtfyTopic);
            actionLayout.Controls.Add(lblNtfyEmail);
            actionLayout.Controls.Add(txtNtfyEmail);
            actionLayout.Controls.Add(btnTestNtfy);
            actionLayout.Dock = DockStyle.Fill;
            actionLayout.Location = new Point(21, 319);
            actionLayout.Margin = new Padding(19, 10, 19, 0);
            actionLayout.Name = "actionLayout";
            actionLayout.Size = new Size(1904, 96);
            actionLayout.TabIndex = 3;
            // 
            // picAccentBar
            // 
            picAccentBar.Dock = DockStyle.Fill;
            picAccentBar.Location = new Point(21, 425);
            picAccentBar.Margin = new Padding(19, 10, 19, 2);
            picAccentBar.Name = "picAccentBar";
            picAccentBar.Size = new Size(1904, 2);
            picAccentBar.SizeMode = PictureBoxSizeMode.StretchImage;
            picAccentBar.TabIndex = 4;
            picAccentBar.TabStop = false;
            // 
            // dashboardLayout
            // 
            dashboardLayout.AutoSize = true;
            dashboardLayout.Controls.Add(lblTotal);
            dashboardLayout.Controls.Add(lblNotTrained);
            dashboardLayout.Controls.Add(lblTraining);
            dashboardLayout.Controls.Add(lblComplete);
            dashboardLayout.Dock = DockStyle.Fill;
            dashboardLayout.Location = new Point(21, 439);
            dashboardLayout.Margin = new Padding(19, 10, 19, 5);
            dashboardLayout.Name = "dashboardLayout";
            dashboardLayout.Size = new Size(1904, 20);
            dashboardLayout.TabIndex = 5;
            // 
            // footerLayout
            // 
            footerLayout.ColumnCount = 5;
            footerLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25F));
            footerLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25F));
            footerLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25F));
            footerLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25F));
            footerLayout.ColumnStyles.Add(new ColumnStyle());
            footerLayout.Controls.Add(pbNotTrained, 0, 0);
            footerLayout.Controls.Add(pbTraining, 1, 0);
            footerLayout.Controls.Add(pbComplete, 2, 0);
            footerLayout.Controls.Add(chkMinimizeTray, 4, 0);
            footerLayout.Dock = DockStyle.Fill;
            footerLayout.Location = new Point(21, 850);
            footerLayout.Margin = new Padding(19, 0, 19, 0);
            footerLayout.Name = "footerLayout";
            footerLayout.RowCount = 1;
            footerLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 120F));
            footerLayout.Size = new Size(1904, 120);
            footerLayout.TabIndex = 7;
            // 
            // MainForm
            // 
            AutoScaleDimensions = new SizeF(11F, 18F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(1946, 984);
            Controls.Add(mainLayout);
            FormBorderStyle = FormBorderStyle.None;
            Margin = new Padding(5, 4, 5, 4);
            MinimumSize = new Size(1666, 816);
            Name = "MainForm";
            StartPosition = FormStartPosition.CenterScreen;
            Text = "Excel Training Monitor";
            ((System.ComponentModel.ISupportInitialize)numScanHours).EndInit();
            ((System.ComponentModel.ISupportInitialize)numScanMinutes).EndInit();
            tabControl1.ResumeLayout(false);
            tabMonitor.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)dgvAlerts).EndInit();
            tabHistory.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)dgvHistory).EndInit();
            tabCharts.ResumeLayout(false);
            chartsLayout.ResumeLayout(false);
            mainLayout.ResumeLayout(false);
            mainLayout.PerformLayout();
            titleBarLayout.ResumeLayout(false);
            titleBarLayout.PerformLayout();
            topLayout.ResumeLayout(false);
            topLayout.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)picThemeLogo).EndInit();
            intervalLayout.ResumeLayout(false);
            intervalLayout.PerformLayout();
            reminderLayout.ResumeLayout(false);
            reminderLayout.PerformLayout();
            themeLayout.ResumeLayout(false);
            themeLayout.PerformLayout();
            fileSearchLayout.ResumeLayout(false);
            fileSearchLayout.PerformLayout();
            actionLayout.ResumeLayout(false);
            actionLayout.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)picAccentBar).EndInit();
            dashboardLayout.ResumeLayout(false);
            dashboardLayout.PerformLayout();
            footerLayout.ResumeLayout(false);
            footerLayout.PerformLayout();
            ResumeLayout(false);
        }

        private DataGridViewTextBoxColumn dataGridViewTextBoxColumn1;
        private DataGridViewTextBoxColumn dataGridViewTextBoxColumn2;
        private DataGridViewTextBoxColumn dataGridViewTextBoxColumn3;
        private DataGridViewTextBoxColumn dataGridViewTextBoxColumn4;
        private DataGridViewTextBoxColumn dataGridViewTextBoxColumn5;
    }
}
