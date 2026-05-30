using System.Drawing;
using System.Windows.Forms;

namespace ExcelTrainingMonitor
{
    partial class MainForm
    {
        private System.ComponentModel.IContainer components = null;
        private System.Windows.Forms.Button btnBrowse;
        private System.Windows.Forms.Button btnStart;
        private System.Windows.Forms.Button btnStop;
        private System.Windows.Forms.Label lblFile;
        private System.Windows.Forms.Label lblInterval;
        private System.Windows.Forms.Label lblReminder;
        private System.Windows.Forms.Label lblTotal;
        private System.Windows.Forms.Label lblNotTrained;
        private System.Windows.Forms.Label lblTraining;
        private System.Windows.Forms.Label lblComplete;
        private System.Windows.Forms.NumericUpDown numInterval;
        private System.Windows.Forms.NumericUpDown numReminderHours;
        private System.Windows.Forms.TextBox txtSearch;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabMonitor;
        private System.Windows.Forms.TabPage tabHistory;
        private System.Windows.Forms.DataGridView dgvAlerts;
        private System.Windows.Forms.DataGridView dgvHistory;
        private System.Windows.Forms.NotifyIcon notifyIcon1;
        private System.Windows.Forms.CheckBox chkMinimizeTray;
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
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

            btnBrowse = new Button();
            btnStart = new Button();
            btnStop = new Button();

            lblFile = new Label();
            lblInterval = new Label();
            lblReminder = new Label();

            lblTotal = new Label();
            lblNotTrained = new Label();
            lblTraining = new Label();
            lblComplete = new Label();

            numInterval = new NumericUpDown();
            numReminderHours = new NumericUpDown();

            txtSearch = new TextBox();

            tabControl1 = new TabControl();

            tabMonitor = new TabPage();
            tabHistory = new TabPage();
            tabMonitor.Name = "tabMonitor";
            tabMonitor.Padding = new Padding(3);
            tabHistory.Name = "tabHistory";
            tabHistory.Padding = new Padding(3);
            tabMonitor.SuspendLayout();
            tabHistory.SuspendLayout();

            dgvAlerts = new DataGridView();
            dgvHistory = new DataGridView();
            dgvHistory.Columns.Add("Employee", "Employee");
            dgvHistory.Columns.Add("Category", "Category");
            dgvHistory.Columns.Add("OldStatus", "Old Status");
            dgvHistory.Columns.Add("NewStatus", "New Status");
            dgvHistory.Columns.Add("Timestamp", "Time");

            notifyIcon1 = new NotifyIcon(components);

            chkMinimizeTray = new CheckBox();
            ((System.ComponentModel.ISupportInitialize)numInterval).BeginInit();
            ((System.ComponentModel.ISupportInitialize)numReminderHours).BeginInit();
            ((System.ComponentModel.ISupportInitialize)dgvAlerts).BeginInit();
            ((System.ComponentModel.ISupportInitialize)dgvHistory).BeginInit();
            // 
            // btnBrowse
            // 
            btnBrowse.Location = new Point(31, 24);
            btnBrowse.Margin = new Padding(5, 4, 5, 4);
            btnBrowse.Name = "btnBrowse";
            btnBrowse.Size = new Size(189, 38);
            btnBrowse.TabIndex = 0;
            btnBrowse.Text = "Browse Excel";
            btnBrowse.UseVisualStyleBackColor = true;
            btnBrowse.Click += btnBrowse_Click;
            // 
            // btnStart
            // 
            btnStart.Location = new Point(251, 24);
            btnStart.Margin = new Padding(5, 4, 5, 4);
            btnStart.Name = "btnStart";
            btnStart.Size = new Size(189, 38);
            btnStart.TabIndex = 1;
            btnStart.Text = "Start";
            btnStart.UseVisualStyleBackColor = true;
            btnStart.Click += btnStart_Click;
            // 
            // btnStop
            // 
            btnStop.Location = new Point(471, 24);
            btnStop.Margin = new Padding(5, 4, 5, 4);
            btnStop.Name = "btnStop";
            btnStop.Size = new Size(189, 38);
            btnStop.TabIndex = 2;
            btnStop.Text = "Stop";
            btnStop.UseVisualStyleBackColor = true;
            btnStop.Click += btnStop_Click;
            // 
            // lblFile
            // 
            lblFile.Location = new Point(31, 84);
            lblFile.Margin = new Padding(5, 0, 5, 0);
            lblFile.Name = "lblFile";
            lblFile.Size = new Size(1414, 30);
            lblFile.TabIndex = 3;
            lblFile.Text = "No Excel file selected";
            // 
            // lblInterval
            // 
            lblInterval.Location = new Point(707, 12);
            lblInterval.Margin = new Padding(5, 0, 5, 0);
            lblInterval.Name = "lblInterval";
            lblInterval.Size = new Size(236, 24);
            lblInterval.TabIndex = 4;
            lblInterval.Text = "Scan Interval (mins)";
            // 
            // lblReminder
            // 
            lblReminder.Location = new Point(974, 12);
            lblReminder.Margin = new Padding(5, 0, 5, 0);
            lblReminder.Name = "lblReminder";
            lblReminder.Size = new Size(236, 24);
            lblReminder.TabIndex = 6;
            lblReminder.Text = "Reminder Hours";
            // 
            // numInterval
            // 
            numInterval.Location = new Point(707, 36);
            numInterval.Margin = new Padding(5, 4, 5, 4);
            numInterval.Maximum = new decimal(new int[] { 1440, 0, 0, 0 });
            numInterval.Minimum = new decimal(new int[] { 1, 0, 0, 0 });
            numInterval.Name = "numInterval";
            numInterval.Size = new Size(189, 31);
            numInterval.TabIndex = 5;
            numInterval.Value = new decimal(new int[] { 5, 0, 0, 0 });
            // 
            // numReminderHours
            // 
            numReminderHours.Location = new Point(974, 36);
            numReminderHours.Margin = new Padding(5, 4, 5, 4);
            numReminderHours.Maximum = new decimal(new int[] { 720, 0, 0, 0 });
            numReminderHours.Minimum = new decimal(new int[] { 1, 0, 0, 0 });
            numReminderHours.Name = "numReminderHours";
            numReminderHours.Size = new Size(189, 31);
            numReminderHours.TabIndex = 7;
            numReminderHours.Value = new decimal(new int[] { 24, 0, 0, 0 });
            // 
            // dgvAlerts
            // 
            dgvAlerts.AllowUserToAddRows = false;
            dgvAlerts.AllowUserToDeleteRows = false;
            dgvAlerts.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dgvAlerts.BackgroundColor = Color.FromArgb(30, 30, 30);
            dgvAlerts.BorderStyle = BorderStyle.None;

            dataGridViewCellStyle1.Alignment = DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = Color.FromArgb(45, 45, 45);
            dataGridViewCellStyle1.Font = new Font("Segoe UI", 10F);
            dataGridViewCellStyle1.ForeColor = Color.White;
            dataGridViewCellStyle1.SelectionBackColor = SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = DataGridViewTriState.True;

            dgvAlerts.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
            dgvAlerts.ColumnHeadersHeight = 34;

            dataGridViewCellStyle2.Alignment = DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.BackColor = Color.FromArgb(37, 37, 38);
            dataGridViewCellStyle2.Font = new Font("Segoe UI", 10F);
            dataGridViewCellStyle2.ForeColor = Color.White;
            dataGridViewCellStyle2.SelectionBackColor = Color.FromArgb(70, 70, 70);
            dataGridViewCellStyle2.SelectionForeColor = Color.White;
            dataGridViewCellStyle2.WrapMode = DataGridViewTriState.False;

            dgvAlerts.DefaultCellStyle = dataGridViewCellStyle2;
            dgvAlerts.EnableHeadersVisualStyles = false;
            dgvAlerts.Font = new Font("Segoe UI", 10F);
            dgvAlerts.Dock = DockStyle.Fill;
            dgvAlerts.Margin = new Padding(5, 4, 5, 4);
            dgvAlerts.Name = "dgvAlerts";
            dgvAlerts.ReadOnly = true;
            dgvAlerts.RowHeadersVisible = false;
            dgvAlerts.RowHeadersWidth = 62;
            dgvAlerts.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dgvAlerts.TabIndex = 8;
            //
            // dgvHistory
            //
            dgvHistory.Dock = DockStyle.Fill;
            dgvHistory.ReadOnly = true;
            dgvHistory.AllowUserToAddRows = false;
            dgvHistory.AllowUserToDeleteRows = false;
            dgvHistory.RowHeadersVisible = false;
            dgvHistory.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

            tabHistory.Controls.Add(dgvHistory);
            // 
            // notifyIcon1
            // 
            notifyIcon1.Text = "Excel Training Monitor";
            notifyIcon1.Visible = true;
            notifyIcon1.DoubleClick += notifyIcon1_DoubleClick;
            // 
            // chkMinimizeTray
            // 
            chkMinimizeTray.Checked = true;
            chkMinimizeTray.CheckState = CheckState.Checked;
            chkMinimizeTray.Location = new Point(31, 744);
            chkMinimizeTray.Margin = new Padding(5, 4, 5, 4);
            chkMinimizeTray.Name = "chkMinimizeTray";
            chkMinimizeTray.Size = new Size(314, 29);
            chkMinimizeTray.TabIndex = 9;
            chkMinimizeTray.Text = "Minimize to tray";
            // 
            // txtSearch
            // 
            txtSearch.Location = new Point(1200, 80);
            txtSearch.Width = 250;
            txtSearch.Name = "txtSearch";
            txtSearch.Size = new Size(150, 31);
            txtSearch.TabIndex = 0;

            txtSearch.TextChanged += TxtSearch_TextChanged;
            // 
            // tabControl1
            // 
            tabControl1.Location = new Point(20, 130);
            tabControl1.Size = new Size(1520, 560);
            tabControl1.Controls.Add(tabMonitor);
            tabControl1.Controls.Add(tabHistory);

            tabMonitor.Text = "Monitor";
            tabHistory.Text = "History";

            lblTotal.Location = new Point(20, 100);
            lblTotal.Size = new Size(120, 25);
            lblTotal.Text = "Total: 0";

            lblNotTrained.Location = new Point(160, 100);
            lblNotTrained.Size = new Size(150, 25);
            lblNotTrained.Text = "Not Trained: 0";

            lblTraining.Location = new Point(330, 100);
            lblTraining.Size = new Size(150, 25);
            lblTraining.Text = "In Training: 0";

            lblComplete.Location = new Point(500, 100);
            lblComplete.Size = new Size(150, 25);
            lblComplete.Text = "Complete: 0";

            lblTotal.Font = new Font("Segoe UI", 10, FontStyle.Bold);
            lblNotTrained.ForeColor = Color.Red;
            lblTraining.ForeColor = Color.Goldenrod;
            lblComplete.ForeColor = Color.LimeGreen;
            // 
            // MainForm
            // 
            AutoScaleDimensions = new SizeF(11F, 18F);
            AutoScaleMode = AutoScaleMode.Font;

            ClientSize = new Size(1571, 816);

            Controls.Add(tabControl1);
            Controls.Add(txtSearch);
            Controls.Add(btnBrowse);
            Controls.Add(btnStart);
            Controls.Add(btnStop);
            Controls.Add(lblFile);
            Controls.Add(lblInterval);
            Controls.Add(numInterval);
            Controls.Add(lblReminder);
            Controls.Add(numReminderHours);
            Controls.Add(lblTotal);
            Controls.Add(lblNotTrained);
            Controls.Add(lblTraining);
            Controls.Add(lblComplete);

            tabMonitor.Controls.Add(dgvAlerts);

            Controls.Add(chkMinimizeTray);

            Margin = new Padding(5, 4, 5, 4);
            Name = "MainForm";
            Text = "Excel Training Monitor";

            ((System.ComponentModel.ISupportInitialize)numInterval).EndInit();
            ((System.ComponentModel.ISupportInitialize)numReminderHours).EndInit();
            ((System.ComponentModel.ISupportInitialize)dgvAlerts).EndInit();
            ((System.ComponentModel.ISupportInitialize)dgvHistory).EndInit();

            tabControl1.ResumeLayout(false);
            tabMonitor.ResumeLayout(false);
            tabHistory.ResumeLayout(false);

            ResumeLayout(false);
            PerformLayout();
        }
    }
}