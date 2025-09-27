using System;
using System.Reflection;
using System.Windows.Forms;

namespace ExcelToOracleImporter
{
    public partial class AboutForm : Form
    {
        public AboutForm()
        {
            InitializeComponent();
            LoadVersionInfo();
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();

            // Form properties
            this.Text = "About - Excel to Oracle Database Importer";
            this.Size = new System.Drawing.Size(450, 350);
            this.StartPosition = FormStartPosition.CenterParent;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;

            // Title Label
            var lblTitle = new Label
            {
                Text = "Excel to Oracle Database Importer",
                Location = new System.Drawing.Point(20, 20),
                Size = new System.Drawing.Size(400, 25),
                Font = new System.Drawing.Font("Microsoft Sans Serif", 14, System.Drawing.FontStyle.Bold),
                TextAlign = System.Drawing.ContentAlignment.MiddleCenter
            };
            this.Controls.Add(lblTitle);

            // Version Label
            var lblVersion = new Label
            {
                Text = "Version: 2.1.0",
                Location = new System.Drawing.Point(20, 60),
                Size = new System.Drawing.Size(400, 20),
                Font = new System.Drawing.Font("Microsoft Sans Serif", 10),
                TextAlign = System.Drawing.ContentAlignment.MiddleCenter
            };
            this.Controls.Add(lblVersion);

            // Author Label
            var lblAuthor = new Label
            {
                Text = "Author: khuongnv@live.com",
                Location = new System.Drawing.Point(20, 90),
                Size = new System.Drawing.Size(400, 20),
                Font = new System.Drawing.Font("Microsoft Sans Serif", 10),
                TextAlign = System.Drawing.ContentAlignment.MiddleCenter
            };
            this.Controls.Add(lblAuthor);

            // Company Label
            var lblCompany = new Label
            {
                Text = "Company: VNPT",
                Location = new System.Drawing.Point(20, 120),
                Size = new System.Drawing.Size(400, 20),
                Font = new System.Drawing.Font("Microsoft Sans Serif", 10),
                TextAlign = System.Drawing.ContentAlignment.MiddleCenter
            };
            this.Controls.Add(lblCompany);

            // Copyright Label
            var lblCopyright = new Label
            {
                Text = "Copyright © 2025 khuongnv@live.com",
                Location = new System.Drawing.Point(20, 150),
                Size = new System.Drawing.Size(400, 20),
                Font = new System.Drawing.Font("Microsoft Sans Serif", 9),
                TextAlign = System.Drawing.ContentAlignment.MiddleCenter
            };
            this.Controls.Add(lblCopyright);

            // Description Label
            var lblDescription = new Label
            {
                Text = "Import Excel data to Oracle Database with:\n• JSON configuration\n• File logging\n• True batch insert\n• Smart table checking",
                Location = new System.Drawing.Point(20, 180),
                Size = new System.Drawing.Size(400, 80),
                Font = new System.Drawing.Font("Microsoft Sans Serif", 9),
                TextAlign = System.Drawing.ContentAlignment.MiddleCenter
            };
            this.Controls.Add(lblDescription);

            // OK Button
            var btnOK = new Button
            {
                Text = "OK",
                Location = new System.Drawing.Point(185, 280),
                Size = new System.Drawing.Size(80, 30),
                DialogResult = DialogResult.OK
            };
            btnOK.Click += (s, e) => this.Close();
            this.Controls.Add(btnOK);

            this.ResumeLayout(false);
        }

        private void LoadVersionInfo()
        {
            try
            {
                var assembly = Assembly.GetExecutingAssembly();
                var version = assembly.GetName().Version;
                
                // Update version label with actual version from assembly
                foreach (Control control in this.Controls)
                {
                    if (control is Label label && label.Text.StartsWith("Version:"))
                    {
                        label.Text = $"Version: {version?.ToString() ?? "2.1.0"}";
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                // Log error but don't crash
                System.Diagnostics.Debug.WriteLine($"Error loading version info: {ex.Message}");
            }
        }
    }
}
