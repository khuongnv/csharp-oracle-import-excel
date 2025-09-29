using System;
using System.Drawing;
using System.Windows.Forms;

namespace ExcelToOracleImporter
{
    public partial class LogForm : Form
    {
        private RichTextBox txtLog;
        private Button btnClose;

        public LogForm()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();

            // Form properties
            this.Text = $"Changelog - Excel to Oracle Importer {VersionInfo.GetFullVersion()}";
            this.Size = new System.Drawing.Size(900, 700);
            this.StartPosition = FormStartPosition.CenterParent;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            // Log TextBox
            txtLog = new RichTextBox
            {
                Location = new System.Drawing.Point(10, 10),
                Size = new System.Drawing.Size(860, 620),
                ReadOnly = true,
                Font = new System.Drawing.Font("Consolas", 9),
                Text = VersionInfo.GetChangelog()
            };
            this.Controls.Add(txtLog);

            // Close Button
            btnClose = new Button
            {
                Text = "Đóng",
                Location = new System.Drawing.Point(800, 640),
                Size = new System.Drawing.Size(70, 30)
            };
            btnClose.Click += (s, e) => this.Close();
            this.Controls.Add(btnClose);

            this.ResumeLayout(false);
        }

    }
}
