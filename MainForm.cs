using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Oracle.ManagedDataAccess.Client;
using OfficeOpenXml;

namespace ExcelToOracleImporter
{
    public partial class MainForm : Form
    {
        private MenuStrip menuStrip;
        private ImportTab importTab;
        private ConnectionManagementTab connectionTab;
        private AppConfig config;

        public MainForm()
        {
            InitializeComponent();
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            
            // Load configuration
            config = AppConfig.Load();
            
            // Initialize tabs
            InitializeTabs();
            
            // Log application start
            FileLogger.LogInfo("Excel to Oracle Importer started");
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();

            // Form properties
            this.Text = $"Excel to Oracle Database Importer {VersionInfo.GetFullVersion()}";
            this.Size = new System.Drawing.Size(1000, 700);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.MaximizeBox = false;

            // Menu Strip
            menuStrip = new MenuStrip();
            
            // Import Excel Menu
            var importMenu = new ToolStripMenuItem("Import Excel");
            importMenu.Click += (s, e) => ShowImportTab();
            menuStrip.Items.Add(importMenu);
            
            // Connection Management Menu
            var connectionMenu = new ToolStripMenuItem("Connection Management");
            connectionMenu.Click += (s, e) => ShowConnectionTab();
            menuStrip.Items.Add(connectionMenu);
            
            // Help Menu (đặt cuối cùng)
            var helpMenu = new ToolStripMenuItem("Help");
            var logMenuItem = new ToolStripMenuItem("Log", null, (s, e) => ShowLogDialog());
            var aboutMenuItem = new ToolStripMenuItem("About", null, (s, e) => ShowAboutDialog());
            helpMenu.DropDownItems.Add(logMenuItem);
            helpMenu.DropDownItems.Add(new ToolStripSeparator());
            helpMenu.DropDownItems.Add(aboutMenuItem);
            menuStrip.Items.Add(helpMenu);
            
            this.Controls.Add(menuStrip);
            this.MainMenuStrip = menuStrip;

            this.ResumeLayout(false);
        }

        private void InitializeTabs()
        {
            // Create Import Tab
            importTab = new ImportTab();
            importTab.LogMessageRequested += (sender, message) => FileLogger.LogInfo(message);
            importTab.StatusUpdateRequested += (sender, status) => { /* Handle status update */ };
            importTab.ConfigurationSaveRequested += (sender, cfg) => { 
                this.config = cfg; 
                config.Save(); 
                FileLogger.LogInfo("Configuration saved from ImportTab");
            };
            importTab.ConfigurationLoadRequested += (sender, cfg) => { /* Handle config load */ };

            // Create Connection Management Tab
            connectionTab = new ConnectionManagementTab();
            connectionTab.LogMessageRequested += (sender, message) => FileLogger.LogInfo(message);
            connectionTab.ConfigurationSaveRequested += (sender, cfg) => { 
                this.config = cfg; 
                config.Save();
                FileLogger.LogInfo("Configuration saved from ConnectionManagementTab");
                // Refresh connection list in import tab when connection is saved
                RefreshImportTabConnections();
            };
            connectionTab.ConfigurationLoadRequested += (sender, cfg) => { /* Handle config load */ };

            // Add UserControls to form
            importTab.Location = new System.Drawing.Point(0, 24);
            importTab.Size = new System.Drawing.Size(1000, 676);
            importTab.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            this.Controls.Add(importTab);

            connectionTab.Location = new System.Drawing.Point(0, 24);
            connectionTab.Size = new System.Drawing.Size(1000, 676);
            connectionTab.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            connectionTab.Visible = false; // Ẩn ban đầu
            this.Controls.Add(connectionTab);

            // Pass shared config to tabs
            importTab.SetConfig(config);
            connectionTab.SetConfig(config);

            // Refresh connection list in import tab
            RefreshImportTabConnections();
        }

        private void RefreshImportTabConnections()
        {
            if (importTab != null && config != null)
            {
                importTab.RefreshConnectionComboBox(config.ConnectionStrings);
            }
        }

        private void ShowImportTab()
        {
            importTab.Visible = true;
            connectionTab.Visible = false;
        }

        private void ShowConnectionTab()
        {
            importTab.Visible = false;
            connectionTab.Visible = true;
        }

        private void ShowLogDialog()
        {
            var logForm = new LogForm();
            logForm.ShowDialog();
        }

        private void ShowAboutDialog()
        {
            var aboutForm = new AboutForm();
            aboutForm.ShowDialog();
        }
    }
}
