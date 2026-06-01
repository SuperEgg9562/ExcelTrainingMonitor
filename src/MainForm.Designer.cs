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
            var dataGridViewCellStyle1 = new DataGridViewCellStyle();
            var dataGridViewCellStyle2 = new DataGridViewCellStyle();

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
            tabHistory = new TabPage();
            tabCharts = new TabPage();
            dgvAlerts = new DataGridView();
            dgvHistory = new DataGridView();
            notifyIcon1 = new NotifyIcon(components);
            chkMinimizeTray = new GlossyCheckBox();
            pbNotTrained = new NeonProgressBar();
            pbTraining = new NeonProgressBar();
            pbComplete = new NeonProgressBar();
            mainLayout = new TableLayoutPanel();
            topLayout = new TableLayoutPanel();
            fileSearchLayout = new TableLayoutPanel();
            actionLayout = new FlowLayoutPanel();
            dashboardLayout = new FlowLayoutPanel();
            footerLayout = new TableLayoutPanel();
            intervalLayout = new TableLayoutPanel();
            reminderLayout = new TableLayoutPanel();
            themeLayout = new TableLayoutPanel();
            chartsLayout = new TableLayoutPanel();
            statusPieChart = new PieChartPanel();
            openPieChart = new PieChartPanel();
            picThemeLogo = new PictureBox();
            picAccentBar = new PictureBox();
            titleBarLayout = new TableLayoutPanel();
            lblWindowTitle = new Label();
            btnWindowMinimize = new GlossyButton();
            btnWindowMaximize = new GlossyButton();
            btnWindowClose = new GlossyButton();

            ((System.ComponentModel.ISupportInitialize)numScanHours).BeginInit();
            ((System.ComponentModel.ISupportInitialize)numScanMinutes).BeginInit();
            ((System.ComponentModel.ISupportInitialize)dgvAlerts).BeginInit();
            ((System.ComponentModel.ISupportInitialize)dgvHistory).BeginInit();
            ((System.ComponentModel.ISupportInitialize)picThemeLogo).BeginInit();
            ((System.ComponentModel.ISupportInitialize)picAccentBar).BeginInit();
            mainLayout.SuspendLayout();
            topLayout.SuspendLayout();
            fileSearchLayout.SuspendLayout();
            actionLayout.SuspendLayout();
            dashboardLayout.SuspendLayout();
            footerLayout.SuspendLayout();
            intervalLayout.SuspendLayout();
            reminderLayout.SuspendLayout();
            themeLayout.SuspendLayout();
            tabControl1.SuspendLayout();
            tabMonitor.SuspendLayout();
            tabHistory.SuspendLayout();
            tabCharts.SuspendLayout();
            chartsLayout.SuspendLayout();
            titleBarLayout.SuspendLayout();
            SuspendLayout();

            lblWindowTitle.AutoSize = true;
            lblWindowTitle.Dock = DockStyle.Fill;
            lblWindowTitle.Font = new Font("Consolas", 11F, FontStyle.Bold);
            lblWindowTitle.Margin = new Padding(10, 0, 0, 0);
            lblWindowTitle.Name = "lblWindowTitle";
            lblWindowTitle.Text = "Excel Training Monitor";
            lblWindowTitle.TextAlign = ContentAlignment.MiddleLeft;
            lblWindowTitle.MouseDown += TitleBar_MouseDown;

            btnWindowMinimize.Dock = DockStyle.Fill;
            btnWindowMinimize.Margin = new Padding(4, 3, 0, 3);
            btnWindowMinimize.Name = "btnWindowMinimize";
            btnWindowMinimize.Text = "_";
            btnWindowMinimize.Click += btnWindowMinimize_Click;

            btnWindowMaximize.Dock = DockStyle.Fill;
            btnWindowMaximize.Margin = new Padding(4, 3, 0, 3);
            btnWindowMaximize.Name = "btnWindowMaximize";
            btnWindowMaximize.Text = "□";
            btnWindowMaximize.Click += btnWindowMaximize_Click;

            btnWindowClose.Dock = DockStyle.Fill;
            btnWindowClose.Margin = new Padding(4, 3, 4, 3);
            btnWindowClose.Name = "btnWindowClose";
            btnWindowClose.Text = "X";
            btnWindowClose.Click += btnWindowClose_Click;

            titleBarLayout.ColumnCount = 4;
            titleBarLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            titleBarLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 44F));
            titleBarLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 44F));
            titleBarLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 44F));
            titleBarLayout.Controls.Add(lblWindowTitle, 0, 0);
            titleBarLayout.Controls.Add(btnWindowMinimize, 1, 0);
            titleBarLayout.Controls.Add(btnWindowMaximize, 2, 0);
            titleBarLayout.Controls.Add(btnWindowClose, 3, 0);
            titleBarLayout.Dock = DockStyle.Fill;
            titleBarLayout.Margin = new Padding(0);
            titleBarLayout.Name = "titleBarLayout";
            titleBarLayout.RowCount = 1;
            titleBarLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            titleBarLayout.MouseDown += TitleBar_MouseDown;

            btnBrowse.AutoSize = true;
            btnBrowse.Dock = DockStyle.Fill;
            btnBrowse.Margin = new Padding(0, 0, 12, 0);
            btnBrowse.MinimumSize = new Size(140, 38);
            btnBrowse.Name = "btnBrowse";
            btnBrowse.TabIndex = 0;
            btnBrowse.Text = "Browse Excel";
            btnBrowse.UseVisualStyleBackColor = true;
            btnBrowse.Click += btnBrowse_Click;

            btnStart.Dock = DockStyle.Fill;
            btnStart.Margin = new Padding(0, 0, 12, 0);
            btnStart.MinimumSize = new Size(100, 38);
            btnStart.Name = "btnStart";
            btnStart.TabIndex = 1;
            btnStart.Text = "Start";
            btnStart.UseVisualStyleBackColor = true;
            btnStart.Click += btnStart_Click;

            btnStop.Dock = DockStyle.Fill;
            btnStop.Margin = new Padding(0, 0, 12, 0);
            btnStop.MinimumSize = new Size(100, 38);
            btnStop.Name = "btnStop";
            btnStop.TabIndex = 2;
            btnStop.Text = "Stop";
            btnStop.UseVisualStyleBackColor = true;
            btnStop.Click += btnStop_Click;

            btnNewExcel.AutoSize = true;
            btnNewExcel.Margin = new Padding(0, 0, 8, 6);
            btnNewExcel.MinimumSize = new Size(104, 34);
            btnNewExcel.Name = "btnNewExcel";
            btnNewExcel.Text = "New Excel";
            btnNewExcel.Click += btnNewExcel_Click;

            btnSaveExcel.AutoSize = true;
            btnSaveExcel.Margin = new Padding(0, 0, 8, 6);
            btnSaveExcel.MinimumSize = new Size(112, 34);
            btnSaveExcel.Name = "btnSaveExcel";
            btnSaveExcel.Text = "Save Edits";
            btnSaveExcel.Click += btnSaveExcel_Click;

            btnExportExcel.AutoSize = true;
            btnExportExcel.Margin = new Padding(0, 0, 8, 6);
            btnExportExcel.MinimumSize = new Size(118, 34);
            btnExportExcel.Name = "btnExportExcel";
            btnExportExcel.Text = "Export Excel";
            btnExportExcel.Click += btnExportExcel_Click;

            btnExportCsv.AutoSize = true;
            btnExportCsv.Margin = new Padding(0, 0, 8, 6);
            btnExportCsv.MinimumSize = new Size(104, 34);
            btnExportCsv.Name = "btnExportCsv";
            btnExportCsv.Text = "Export CSV";
            btnExportCsv.Click += btnExportCsv_Click;

            btnPrintPdf.AutoSize = true;
            btnPrintPdf.Margin = new Padding(0, 0, 12, 6);
            btnPrintPdf.MinimumSize = new Size(104, 34);
            btnPrintPdf.Name = "btnPrintPdf";
            btnPrintPdf.Text = "Print PDF";
            btnPrintPdf.Click += btnPrintPdf_Click;

            lblInterval.AutoSize = true;
            lblInterval.Dock = DockStyle.Fill;
            lblInterval.Name = "lblInterval";
            lblInterval.Text = "Scan Interval";
            lblInterval.TextAlign = ContentAlignment.BottomLeft;

            lblScanHours.AutoSize = true;
            lblScanHours.Dock = DockStyle.Fill;
            lblScanHours.Name = "lblScanHours";
            lblScanHours.Text = "Hours";

            lblScanMinutes.AutoSize = true;
            lblScanMinutes.Dock = DockStyle.Fill;
            lblScanMinutes.Name = "lblScanMinutes";
            lblScanMinutes.Text = "Minutes";

            numScanHours.Dock = DockStyle.Fill;
            numScanHours.Maximum = new decimal(new int[] { 24, 0, 0, 0 });
            numScanHours.Name = "numScanHours";
            numScanHours.TabIndex = 5;

            numScanMinutes.Dock = DockStyle.Fill;
            numScanMinutes.Maximum = new decimal(new int[] { 59, 0, 0, 0 });
            numScanMinutes.Name = "numScanMinutes";
            numScanMinutes.TabIndex = 6;
            numScanMinutes.Value = new decimal(new int[] { 5, 0, 0, 0 });

            intervalLayout.ColumnCount = 2;
            intervalLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
            intervalLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
            intervalLayout.Controls.Add(lblInterval, 0, 0);
            intervalLayout.SetColumnSpan(lblInterval, 2);
            intervalLayout.Controls.Add(lblScanHours, 0, 1);
            intervalLayout.Controls.Add(lblScanMinutes, 1, 1);
            intervalLayout.Controls.Add(numScanHours, 0, 2);
            intervalLayout.Controls.Add(numScanMinutes, 1, 2);
            intervalLayout.Dock = DockStyle.Fill;
            intervalLayout.Margin = new Padding(0);
            intervalLayout.RowCount = 3;
            intervalLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            intervalLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            intervalLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            lblReminder.AutoSize = true;
            lblReminder.Dock = DockStyle.Fill;
            lblReminder.Name = "lblReminder";
            lblReminder.Text = "Reminder Date";
            lblReminder.TextAlign = ContentAlignment.BottomLeft;

            chkReminderEnabled.AutoSize = true;
            chkReminderEnabled.Dock = DockStyle.Fill;
            chkReminderEnabled.Name = "chkReminderEnabled";
            chkReminderEnabled.Text = "Enable reminder";

            chkNtfyEnabled.AutoSize = true;
            chkNtfyEnabled.Margin = new Padding(0, 6, 10, 6);
            chkNtfyEnabled.Name = "chkNtfyEnabled";
            chkNtfyEnabled.Text = "ntfy.sh";

            dtpReminderDate.CustomFormat = "yyyy-MM-dd HH:mm";
            dtpReminderDate.Dock = DockStyle.Fill;
            dtpReminderDate.Format = DateTimePickerFormat.Custom;
            dtpReminderDate.Name = "dtpReminderDate";

            reminderLayout.ColumnCount = 1;
            reminderLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            reminderLayout.Controls.Add(lblReminder, 0, 0);
            reminderLayout.Controls.Add(dtpReminderDate, 0, 1);
            reminderLayout.Controls.Add(chkReminderEnabled, 0, 2);
            reminderLayout.Dock = DockStyle.Fill;
            reminderLayout.Margin = new Padding(0);
            reminderLayout.RowCount = 3;
            reminderLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            reminderLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            reminderLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            lblTheme.AutoSize = true;
            lblTheme.Dock = DockStyle.Fill;
            lblTheme.Name = "lblTheme";
            lblTheme.Text = "Theme";
            lblTheme.TextAlign = ContentAlignment.BottomLeft;

            cboTheme.Dock = DockStyle.Fill;
            cboTheme.DropDownStyle = ComboBoxStyle.DropDownList;
            cboTheme.Items.AddRange(new object[] { "Dark" });
            cboTheme.Name = "cboTheme";

            themeLayout.ColumnCount = 1;
            themeLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            themeLayout.Controls.Add(lblTheme, 0, 0);
            themeLayout.Controls.Add(cboTheme, 0, 1);
            themeLayout.Dock = DockStyle.Fill;
            themeLayout.Margin = new Padding(0);
            themeLayout.RowCount = 2;
            themeLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            themeLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            topLayout.ColumnCount = 7;
            topLayout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            topLayout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            topLayout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            topLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            topLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 190F));
            topLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 230F));
            topLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 140F));
            topLayout.Controls.Add(btnBrowse, 0, 0);
            topLayout.Controls.Add(btnStart, 1, 0);
            topLayout.Controls.Add(btnStop, 2, 0);
            topLayout.Controls.Add(picThemeLogo, 3, 0);
            topLayout.Controls.Add(intervalLayout, 4, 0);
            topLayout.Controls.Add(reminderLayout, 5, 0);
            topLayout.Controls.Add(themeLayout, 6, 0);
            topLayout.Dock = DockStyle.Fill;
            topLayout.Margin = new Padding(12, 10, 12, 0);
            topLayout.RowCount = 1;
            topLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            picThemeLogo.Dock = DockStyle.Right;
            picThemeLogo.Margin = new Padding(0, 0, 16, 0);
            picThemeLogo.Name = "picThemeLogo";
            picThemeLogo.Size = new Size(54, 54);
            picThemeLogo.SizeMode = PictureBoxSizeMode.Zoom;
            picThemeLogo.TabStop = false;

            lblFile.AutoEllipsis = true;
            lblFile.Dock = DockStyle.Fill;
            lblFile.Margin = new Padding(0, 6, 16, 0);
            lblFile.Name = "lblFile";
            lblFile.TabIndex = 3;
            lblFile.Text = "No Excel file selected";
            lblFile.TextAlign = ContentAlignment.MiddleLeft;

            txtSearch.Dock = DockStyle.Fill;
            txtSearch.Margin = new Padding(0, 6, 0, 0);
            txtSearch.Name = "txtSearch";
            txtSearch.PlaceholderText = "Search";
            txtSearch.TabIndex = 8;
            txtSearch.TextChanged += TxtSearch_TextChanged;

            fileSearchLayout.ColumnCount = 2;
            fileSearchLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            fileSearchLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 280F));
            fileSearchLayout.Controls.Add(lblFile, 0, 0);
            fileSearchLayout.Controls.Add(txtSearch, 1, 0);
            fileSearchLayout.Dock = DockStyle.Fill;
            fileSearchLayout.Margin = new Padding(12, 8, 12, 0);
            fileSearchLayout.RowCount = 1;
            fileSearchLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            lblNtfyTopic.AutoSize = true;
            lblNtfyTopic.Margin = new Padding(0, 8, 4, 0);
            lblNtfyTopic.Name = "lblNtfyTopic";
            lblNtfyTopic.Text = "Topic";

            txtNtfyTopic.Margin = new Padding(0, 4, 10, 6);
            txtNtfyTopic.Name = "txtNtfyTopic";
            txtNtfyTopic.Size = new Size(160, 23);

            lblNtfyEmail.AutoSize = true;
            lblNtfyEmail.Margin = new Padding(0, 8, 4, 0);
            lblNtfyEmail.Name = "lblNtfyEmail";
            lblNtfyEmail.Text = "Email";

            txtNtfyEmail.Margin = new Padding(0, 4, 10, 6);
            txtNtfyEmail.Name = "txtNtfyEmail";
            txtNtfyEmail.Size = new Size(190, 23);

            btnTestNtfy.AutoSize = true;
            btnTestNtfy.Margin = new Padding(0, 0, 0, 6);
            btnTestNtfy.MinimumSize = new Size(98, 34);
            btnTestNtfy.Name = "btnTestNtfy";
            btnTestNtfy.Text = "Test ntfy";
            btnTestNtfy.Click += btnTestNtfy_Click;

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
            actionLayout.Margin = new Padding(12, 8, 12, 0);
            actionLayout.Name = "actionLayout";
            actionLayout.WrapContents = true;

            picAccentBar.Dock = DockStyle.Fill;
            picAccentBar.Margin = new Padding(12, 8, 12, 2);
            picAccentBar.Name = "picAccentBar";
            picAccentBar.SizeMode = PictureBoxSizeMode.StretchImage;
            picAccentBar.TabStop = false;

            lblTotal.AutoSize = true;
            lblTotal.Font = new Font("Segoe UI", 10F, FontStyle.Bold);
            lblTotal.Margin = new Padding(0, 0, 24, 0);
            lblTotal.Name = "lblTotal";
            lblTotal.Text = "Total: 0";

            lblNotTrained.AutoSize = true;
            lblNotTrained.ForeColor = Color.FromArgb(255, 70, 70);
            lblNotTrained.Margin = new Padding(0, 0, 24, 0);
            lblNotTrained.Name = "lblNotTrained";
            lblNotTrained.Text = "Not Trained: 0";

            lblTraining.AutoSize = true;
            lblTraining.ForeColor = Color.FromArgb(255, 210, 28);
            lblTraining.Margin = new Padding(0, 0, 24, 0);
            lblTraining.Name = "lblTraining";
            lblTraining.Text = "In Training: 0";

            lblComplete.AutoSize = true;
            lblComplete.ForeColor = Color.FromArgb(0, 255, 40);
            lblComplete.Margin = new Padding(0);
            lblComplete.Name = "lblComplete";
            lblComplete.Text = "Complete: 0";

            dashboardLayout.AutoSize = true;
            dashboardLayout.Controls.Add(lblTotal);
            dashboardLayout.Controls.Add(lblNotTrained);
            dashboardLayout.Controls.Add(lblTraining);
            dashboardLayout.Controls.Add(lblComplete);
            dashboardLayout.Dock = DockStyle.Fill;
            dashboardLayout.Margin = new Padding(12, 8, 12, 4);
            dashboardLayout.WrapContents = true;

            dataGridViewCellStyle1.Alignment = DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = Color.FromArgb(3, 45, 13);
            dataGridViewCellStyle1.Font = new Font("Segoe UI", 10F);
            dataGridViewCellStyle1.ForeColor = Color.FromArgb(0, 255, 40);
            dataGridViewCellStyle1.SelectionBackColor = Color.FromArgb(0, 220, 35);
            dataGridViewCellStyle1.SelectionForeColor = Color.Black;
            dataGridViewCellStyle1.WrapMode = DataGridViewTriState.True;

            dataGridViewCellStyle2.Alignment = DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.BackColor = Color.FromArgb(8, 10, 9);
            dataGridViewCellStyle2.Font = new Font("Segoe UI", 10F);
            dataGridViewCellStyle2.ForeColor = Color.FromArgb(0, 255, 40);
            dataGridViewCellStyle2.SelectionBackColor = Color.FromArgb(0, 220, 35);
            dataGridViewCellStyle2.SelectionForeColor = Color.Black;
            dataGridViewCellStyle2.WrapMode = DataGridViewTriState.False;

            dgvAlerts.AllowUserToAddRows = false;
            dgvAlerts.AllowUserToDeleteRows = false;
            dgvAlerts.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dgvAlerts.BackgroundColor = Color.FromArgb(8, 10, 9);
            dgvAlerts.BorderStyle = BorderStyle.None;
            dgvAlerts.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
            dgvAlerts.ColumnHeadersHeight = 34;
            dgvAlerts.DefaultCellStyle = dataGridViewCellStyle2;
            dgvAlerts.Dock = DockStyle.Fill;
            dgvAlerts.EnableHeadersVisualStyles = false;
            dgvAlerts.Font = new Font("Segoe UI", 10F);
            dgvAlerts.Margin = new Padding(0);
            dgvAlerts.Name = "dgvAlerts";
            dgvAlerts.ReadOnly = true;
            dgvAlerts.RowHeadersVisible = false;
            dgvAlerts.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dgvAlerts.TabIndex = 9;

            dgvHistory.AllowUserToAddRows = false;
            dgvHistory.AllowUserToDeleteRows = false;
            dgvHistory.AutoGenerateColumns = false;
            dgvHistory.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dgvHistory.Dock = DockStyle.Fill;
            dgvHistory.Margin = new Padding(0);
            dgvHistory.Name = "dgvHistory";
            dgvHistory.ReadOnly = true;
            dgvHistory.RowHeadersVisible = false;
            dgvHistory.Columns.Add(new DataGridViewTextBoxColumn
            {
                DataPropertyName = "Employee",
                HeaderText = "Employee",
                Name = "Employee"
            });
            dgvHistory.Columns.Add(new DataGridViewTextBoxColumn
            {
                DataPropertyName = "Category",
                HeaderText = "Category",
                Name = "Category"
            });
            dgvHistory.Columns.Add(new DataGridViewTextBoxColumn
            {
                DataPropertyName = "OldStatus",
                HeaderText = "Old Status",
                Name = "OldStatus"
            });
            dgvHistory.Columns.Add(new DataGridViewTextBoxColumn
            {
                DataPropertyName = "NewStatus",
                HeaderText = "New Status",
                Name = "NewStatus"
            });
            dgvHistory.Columns.Add(new DataGridViewTextBoxColumn
            {
                DataPropertyName = "Timestamp",
                HeaderText = "Time",
                Name = "Timestamp"
            });

            tabMonitor.Controls.Add(dgvAlerts);
            tabMonitor.Name = "tabMonitor";
            tabMonitor.Padding = new Padding(3);
            tabMonitor.Text = "Monitor";

            tabHistory.Controls.Add(dgvHistory);
            tabHistory.Name = "tabHistory";
            tabHistory.Padding = new Padding(3);
            tabHistory.Text = "History";

            statusPieChart.ChartTitle = "All Training Statuses";
            statusPieChart.Dock = DockStyle.Fill;
            statusPieChart.Margin = new Padding(0, 0, 8, 0);
            statusPieChart.Name = "statusPieChart";

            openPieChart.ChartTitle = "Open Training Items";
            openPieChart.Dock = DockStyle.Fill;
            openPieChart.Margin = new Padding(8, 0, 0, 0);
            openPieChart.Name = "openPieChart";

            chartsLayout.ColumnCount = 2;
            chartsLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
            chartsLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
            chartsLayout.Controls.Add(statusPieChart, 0, 0);
            chartsLayout.Controls.Add(openPieChart, 1, 0);
            chartsLayout.Dock = DockStyle.Fill;
            chartsLayout.Margin = new Padding(0);
            chartsLayout.RowCount = 1;
            chartsLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            tabCharts.Controls.Add(chartsLayout);
            tabCharts.Name = "tabCharts";
            tabCharts.Padding = new Padding(3);
            tabCharts.Text = "Graphs";

            tabControl1.Controls.Add(tabMonitor);
            tabControl1.Controls.Add(tabHistory);
            tabControl1.Controls.Add(tabCharts);
            tabControl1.Dock = DockStyle.Fill;
            tabControl1.Margin = new Padding(12, 0, 12, 12);
            tabControl1.Name = "tabControl1";
            tabControl1.SelectedIndex = 0;

            pbNotTrained.Dock = DockStyle.Fill;
            pbNotTrained.Margin = new Padding(0, 0, 12, 0);
            pbNotTrained.Name = "pbNotTrained";

            pbTraining.Dock = DockStyle.Fill;
            pbTraining.Margin = new Padding(0, 0, 12, 0);
            pbTraining.Name = "pbTraining";

            pbComplete.Dock = DockStyle.Fill;
            pbComplete.Margin = new Padding(0, 0, 24, 0);
            pbComplete.Name = "pbComplete";

            chkMinimizeTray.AutoSize = true;
            chkMinimizeTray.Checked = true;
            chkMinimizeTray.CheckState = CheckState.Checked;
            chkMinimizeTray.Dock = DockStyle.Right;
            chkMinimizeTray.Margin = new Padding(0);
            chkMinimizeTray.Name = "chkMinimizeTray";
            chkMinimizeTray.TabIndex = 10;
            chkMinimizeTray.Text = "Minimize to tray";

            footerLayout.ColumnCount = 5;
            footerLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25F));
            footerLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25F));
            footerLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25F));
            footerLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25F));
            footerLayout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            footerLayout.Controls.Add(pbNotTrained, 0, 0);
            footerLayout.Controls.Add(pbTraining, 1, 0);
            footerLayout.Controls.Add(pbComplete, 2, 0);
            footerLayout.Controls.Add(chkMinimizeTray, 4, 0);
            footerLayout.Dock = DockStyle.Fill;
            footerLayout.Margin = new Padding(12, 0, 12, 0);
            footerLayout.RowCount = 1;
            footerLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 28F));

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
            mainLayout.Margin = new Padding(0);
            mainLayout.Padding = new Padding(1, 1, 1, 12);
            mainLayout.RowCount = 8;
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 38F));
            mainLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            mainLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            mainLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 12F));
            mainLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            mainLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            notifyIcon1.Text = "Excel Training Monitor";
            notifyIcon1.Visible = true;
            notifyIcon1.DoubleClick += notifyIcon1_DoubleClick;

            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(1280, 820);
            Controls.Add(mainLayout);
            FormBorderStyle = FormBorderStyle.None;
            MinimumSize = new Size(1060, 680);
            Name = "MainForm";
            StartPosition = FormStartPosition.CenterScreen;
            Text = "Excel Training Monitor";

            ((System.ComponentModel.ISupportInitialize)numScanHours).EndInit();
            ((System.ComponentModel.ISupportInitialize)numScanMinutes).EndInit();
            ((System.ComponentModel.ISupportInitialize)dgvAlerts).EndInit();
            ((System.ComponentModel.ISupportInitialize)dgvHistory).EndInit();
            ((System.ComponentModel.ISupportInitialize)picThemeLogo).EndInit();
            ((System.ComponentModel.ISupportInitialize)picAccentBar).EndInit();
            mainLayout.ResumeLayout(false);
            mainLayout.PerformLayout();
            topLayout.ResumeLayout(false);
            topLayout.PerformLayout();
            fileSearchLayout.ResumeLayout(false);
            fileSearchLayout.PerformLayout();
            actionLayout.ResumeLayout(false);
            actionLayout.PerformLayout();
            dashboardLayout.ResumeLayout(false);
            dashboardLayout.PerformLayout();
            footerLayout.ResumeLayout(false);
            footerLayout.PerformLayout();
            intervalLayout.ResumeLayout(false);
            intervalLayout.PerformLayout();
            reminderLayout.ResumeLayout(false);
            reminderLayout.PerformLayout();
            themeLayout.ResumeLayout(false);
            themeLayout.PerformLayout();
            tabControl1.ResumeLayout(false);
            tabMonitor.ResumeLayout(false);
            tabHistory.ResumeLayout(false);
            tabCharts.ResumeLayout(false);
            chartsLayout.ResumeLayout(false);
            titleBarLayout.ResumeLayout(false);
            titleBarLayout.PerformLayout();
            ResumeLayout(false);
        }
    }
}
