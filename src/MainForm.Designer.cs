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

        private System.Windows.Forms.NumericUpDown numInterval;
        private System.Windows.Forms.NumericUpDown numReminderHours;

        private System.Windows.Forms.DataGridView dgvAlerts;

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
            this.components =
                new System.ComponentModel.Container();

            this.btnBrowse =
                new System.Windows.Forms.Button();

            this.btnStart =
                new System.Windows.Forms.Button();

            this.btnStop =
                new System.Windows.Forms.Button();

            this.lblFile =
                new System.Windows.Forms.Label();

            this.lblInterval =
                new System.Windows.Forms.Label();

            this.lblReminder =
                new System.Windows.Forms.Label();

            this.numInterval =
                new System.Windows.Forms.NumericUpDown();

            this.numReminderHours =
                new System.Windows.Forms.NumericUpDown();

            this.dgvAlerts =
                new System.Windows.Forms.DataGridView();

            this.notifyIcon1 =
                new System.Windows.Forms.NotifyIcon(this.components);

            this.chkMinimizeTray =
                new System.Windows.Forms.CheckBox();

            ((System.ComponentModel.ISupportInitialize)
                (this.numInterval)).BeginInit();

            ((System.ComponentModel.ISupportInitialize)
                (this.numReminderHours)).BeginInit();

            ((System.ComponentModel.ISupportInitialize)
                (this.dgvAlerts)).BeginInit();

            this.SuspendLayout();

            this.btnBrowse.Location =
                new System.Drawing.Point(20, 20);

            this.btnBrowse.Name =
                "btnBrowse";

            this.btnBrowse.Size =
                new System.Drawing.Size(120, 32);

            this.btnBrowse.Text =
                "Browse Excel";

            this.btnBrowse.UseVisualStyleBackColor =
                true;

            this.btnBrowse.Click +=
                new System.EventHandler(this.btnBrowse_Click);

            this.btnStart.Location =
                new System.Drawing.Point(160, 20);

            this.btnStart.Name =
                "btnStart";

            this.btnStart.Size =
                new System.Drawing.Size(120, 32);

            this.btnStart.Text =
                "Start";

            this.btnStart.UseVisualStyleBackColor =
                true;

            this.btnStart.Click +=
                new System.EventHandler(this.btnStart_Click);

            this.btnStop.Location =
                new System.Drawing.Point(300, 20);

            this.btnStop.Name =
                "btnStop";

            this.btnStop.Size =
                new System.Drawing.Size(120, 32);

            this.btnStop.Text =
                "Stop";

            this.btnStop.UseVisualStyleBackColor =
                true;

            this.btnStop.Click +=
                new System.EventHandler(this.btnStop_Click);

            this.lblFile.Location =
                new System.Drawing.Point(20, 70);

            this.lblFile.Name =
                "lblFile";

            this.lblFile.Size =
                new System.Drawing.Size(900, 25);

            this.lblFile.Text =
                "No Excel file selected";

            this.lblInterval.Location =
                new System.Drawing.Point(450, 10);

            this.lblInterval.Size =
                new System.Drawing.Size(150, 20);

            this.lblInterval.Text =
                "Scan Interval (mins)";

            this.lblReminder.Location =
                new System.Drawing.Point(620, 10);

            this.lblReminder.Size =
                new System.Drawing.Size(150, 20);

            this.lblReminder.Text =
                "Reminder Hours";

            this.numInterval.Location =
                new System.Drawing.Point(450, 30);

            this.numInterval.Minimum = 1;

            this.numInterval.Maximum = 1440;

            this.numInterval.Value = 5;

            this.numReminderHours.Location =
                new System.Drawing.Point(620, 30);

            this.numReminderHours.Minimum = 1;

            this.numReminderHours.Maximum = 720;

            this.numReminderHours.Value = 24;

            this.dgvAlerts.AutoGenerateColumns =
                false;

            this.dgvAlerts.Location =
                new System.Drawing.Point(12, 110);

            this.dgvAlerts.Size =
                new System.Drawing.Size(960, 500);

            this.dgvAlerts.AllowUserToAddRows =
                false;

            this.dgvAlerts.AllowUserToDeleteRows =
                false;

            this.dgvAlerts.ReadOnly =
                true;

            this.dgvAlerts.RowHeadersVisible =
                false;

            this.dgvAlerts.SelectionMode =
                System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;

            this.dgvAlerts.AutoSizeColumnsMode =
                System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;

            this.dgvAlerts.BackgroundColor =
                System.Drawing.Color.FromArgb(30, 30, 30);

            this.dgvAlerts.BorderStyle =
                System.Windows.Forms.BorderStyle.None;

            this.dgvAlerts.EnableHeadersVisualStyles =
                false;

            this.dgvAlerts.ColumnHeadersDefaultCellStyle.BackColor =
                System.Drawing.Color.FromArgb(45, 45, 45);

            this.dgvAlerts.ColumnHeadersDefaultCellStyle.ForeColor =
                System.Drawing.Color.White;

            this.dgvAlerts.DefaultCellStyle.BackColor =
                System.Drawing.Color.FromArgb(37, 37, 38);

            this.dgvAlerts.DefaultCellStyle.ForeColor =
                System.Drawing.Color.White;

            this.dgvAlerts.DefaultCellStyle.SelectionBackColor =
                System.Drawing.Color.FromArgb(70, 70, 70);

            this.dgvAlerts.DefaultCellStyle.SelectionForeColor =
                System.Drawing.Color.White;

            this.dgvAlerts.Font =
                new System.Drawing.Font("Segoe UI", 10F);

            DataGridViewTextBoxColumn employeeColumn =
                new DataGridViewTextBoxColumn();

            employeeColumn.Name =
                "EmployeeName";

            employeeColumn.HeaderText =
                "Employee";

            employeeColumn.DataPropertyName =
                "EmployeeName";

            employeeColumn.AutoSizeMode =
                DataGridViewAutoSizeColumnMode.Fill;

            DataGridViewTextBoxColumn categoryColumn =
                new DataGridViewTextBoxColumn();

            categoryColumn.Name =
                "Category";

            categoryColumn.HeaderText =
                "Training Category";

            categoryColumn.DataPropertyName =
                "Category";

            categoryColumn.AutoSizeMode =
                DataGridViewAutoSizeColumnMode.Fill;

            DataGridViewTextBoxColumn statusColumn =
                new DataGridViewTextBoxColumn();

            statusColumn.Name =
                "Status";

            statusColumn.HeaderText =
                "Status";

            statusColumn.DataPropertyName =
                "Status";

            statusColumn.AutoSizeMode =
                DataGridViewAutoSizeColumnMode.Fill;

            DataGridViewTextBoxColumn timestampColumn =
                new DataGridViewTextBoxColumn();

            timestampColumn.Name =
                "Timestamp";

            timestampColumn.HeaderText =
                "Detected";

            timestampColumn.DataPropertyName =
                "Timestamp";

            timestampColumn.AutoSizeMode =
                DataGridViewAutoSizeColumnMode.Fill;

            this.dgvAlerts.Columns.Add(employeeColumn);

            this.dgvAlerts.Columns.Add(categoryColumn);

            this.dgvAlerts.Columns.Add(statusColumn);

            this.dgvAlerts.Columns.Add(timestampColumn);

            this.chkMinimizeTray.Location =
                new System.Drawing.Point(20, 620);

            this.chkMinimizeTray.Size =
                new System.Drawing.Size(200, 24);

            this.chkMinimizeTray.Text =
                "Minimize to tray";

            this.chkMinimizeTray.Checked =
                true;

            this.notifyIcon1.Text =
                "Excel Training Monitor";

            this.notifyIcon1.Visible =
                true;

            this.notifyIcon1.DoubleClick +=
                new System.EventHandler(
                    this.notifyIcon1_DoubleClick);

            this.AutoScaleDimensions =
                new System.Drawing.SizeF(7F, 15F);

            this.AutoScaleMode =
                System.Windows.Forms.AutoScaleMode.Font;

            this.ClientSize =
                new System.Drawing.Size(1000, 680);

            this.Controls.Add(this.btnBrowse);

            this.Controls.Add(this.btnStart);

            this.Controls.Add(this.btnStop);

            this.Controls.Add(this.lblFile);

            this.Controls.Add(this.lblInterval);

            this.Controls.Add(this.numInterval);

            this.Controls.Add(this.lblReminder);

            this.Controls.Add(this.numReminderHours);

            this.Controls.Add(this.dgvAlerts);

            this.Controls.Add(this.chkMinimizeTray);

            this.Name =
                "MainForm";

            this.Text =
                "Excel Training Monitor";

            ((System.ComponentModel.ISupportInitialize)
                (this.numInterval)).EndInit();

            ((System.ComponentModel.ISupportInitialize)
                (this.numReminderHours)).EndInit();

            ((System.ComponentModel.ISupportInitialize)
                (this.dgvAlerts)).EndInit();

            this.ResumeLayout(false);
        }
    }
}