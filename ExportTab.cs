using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using Oracle.ManagedDataAccess.Client;
using OfficeOpenXml;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelToOracleImporter
{
    public partial class ExportTab : UserControl
    {
        public event EventHandler<string> LogMessageRequested;
        public event EventHandler<string> StatusUpdateRequested;

        private ComboBox cmbConnectionString;
        private Button btnTestConnection;
        private TextBox txtSqlQuery;
        private Button btnExport;
        private ProgressBar progressBar;
        private Label lblStatus;
        private TextBox txtLog;
        private Button btnOpenLogs;
        private NumericUpDown numMaxRowsPerSheet;
        private TextBox txtExportPath;
        private Button btnSelectExportFolder;

        private AppConfig config;

        public ExportTab()
        {
            InitializeComponent();
            InitializeExportTab();
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();

            // Connection String Selection
            var lblConnectionString = new Label
            {
                Text = "Connection String:",
                Location = new System.Drawing.Point(20, 20),
                Size = new System.Drawing.Size(120, 20)
            };
            this.Controls.Add(lblConnectionString);

            cmbConnectionString = new ComboBox
            {
                Location = new System.Drawing.Point(150, 18),
                Size = new System.Drawing.Size(600, 20),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            cmbConnectionString.SelectedIndexChanged += CmbConnectionString_SelectedIndexChanged;
            this.Controls.Add(cmbConnectionString);

            btnTestConnection = new Button
            {
                Text = "Test Connection",
                Location = new System.Drawing.Point(760, 17),
                Size = new System.Drawing.Size(120, 25)
            };
            btnTestConnection.Click += BtnTestConnection_Click;
            this.Controls.Add(btnTestConnection);

            // SQL Query
            var lblSqlQuery = new Label
            {
                Text = "SQL Query:",
                Location = new System.Drawing.Point(20, 60),
                Size = new System.Drawing.Size(120, 20)
            };
            this.Controls.Add(lblSqlQuery);

            txtSqlQuery = new TextBox
            {
                Location = new System.Drawing.Point(20, 85),
                Size = new System.Drawing.Size(960, 150),
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                Text = "SELECT * FROM your_table_name WHERE rownum <= 1000"
            };
            this.Controls.Add(txtSqlQuery);

            // Max Rows Per Sheet
            var lblMaxRows = new Label
            {
                Text = "Max Rows Per Sheet:",
                Location = new System.Drawing.Point(20, 250),
                Size = new System.Drawing.Size(120, 20)
            };
            this.Controls.Add(lblMaxRows);

            numMaxRowsPerSheet = new NumericUpDown
            {
                Location = new System.Drawing.Point(150, 248),
                Size = new System.Drawing.Size(100, 20),
                Minimum = 1000,
                Maximum = 10000000,
                Value = 1000000,
                Increment = 100000
            };
            this.Controls.Add(numMaxRowsPerSheet);

            // Export Path
            var lblExportPath = new Label
            {
                Text = "Export Folder:",
                Location = new System.Drawing.Point(20, 280),
                Size = new System.Drawing.Size(120, 20)
            };
            this.Controls.Add(lblExportPath);

            txtExportPath = new TextBox
            {
                Location = new System.Drawing.Point(150, 278),
                Size = new System.Drawing.Size(600, 20),
                ReadOnly = true,
                Text = Application.StartupPath
            };
            this.Controls.Add(txtExportPath);

            btnSelectExportFolder = new Button
            {
                Text = "Select Folder",
                Location = new System.Drawing.Point(760, 277),
                Size = new System.Drawing.Size(120, 25)
            };
            btnSelectExportFolder.Click += BtnSelectExportFolder_Click;
            this.Controls.Add(btnSelectExportFolder);

            // Export Button
            btnExport = new Button
            {
                Text = "Export to Excel",
                Location = new System.Drawing.Point(20, 320),
                Size = new System.Drawing.Size(120, 30),
                BackColor = Color.LightBlue
            };
            btnExport.Click += BtnExport_Click;
            this.Controls.Add(btnExport);

            // Open Logs Button
            btnOpenLogs = new Button
            {
                Text = "Open Logs",
                Location = new System.Drawing.Point(150, 320),
                Size = new System.Drawing.Size(100, 30)
            };
            btnOpenLogs.Click += BtnOpenLogs_Click;
            this.Controls.Add(btnOpenLogs);

            // Progress Bar
            progressBar = new ProgressBar
            {
                Location = new System.Drawing.Point(20, 370),
                Size = new System.Drawing.Size(960, 20),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            this.Controls.Add(progressBar);

            // Status Label
            lblStatus = new Label
            {
                Text = "Status: Ready",
                Location = new System.Drawing.Point(20, 400),
                Size = new System.Drawing.Size(960, 20),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            this.Controls.Add(lblStatus);

            // Log Textbox
            txtLog = new TextBox
            {
                Location = new System.Drawing.Point(20, 430),
                Size = new System.Drawing.Size(960, 170),
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Vertical,
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
            };
            this.Controls.Add(txtLog);

            this.ResumeLayout(false);
            this.PerformLayout();
        }

        private void InitializeExportTab()
        {
            LogMessage("Export tab initialized.");
        }

        public void SetConfig(AppConfig appConfig)
        {
            this.config = appConfig;
            LoadConnectionStrings();
        }

        private void LoadConnectionStrings()
        {
            if (config?.ConnectionStrings != null)
            {
                cmbConnectionString.Items.Clear();
                foreach (var conn in config.ConnectionStrings.OrderBy(c => c.Order))
                {
                    cmbConnectionString.Items.Add($"{conn.Name} - {conn.ConnectionString}");
                }
                if (cmbConnectionString.Items.Count > 0)
                {
                    cmbConnectionString.SelectedIndex = 0;
                }
            }
        }

        private void CmbConnectionString_SelectedIndexChanged(object sender, EventArgs e)
        {
            // Connection string selection changed
        }

        private async void BtnTestConnection_Click(object sender, EventArgs e)
        {
            if (cmbConnectionString.SelectedItem == null)
            {
                LogMessage("Please select a connection string first.");
                return;
            }

            btnTestConnection.Enabled = false;
            UpdateStatus("Testing connection...");

            try
            {
                var connectionString = cmbConnectionString.SelectedItem.ToString().Split(" - ")[1];
                
                using (var connection = new OracleConnection(connectionString))
                {
                    await connection.OpenAsync();
                    LogMessage("✅ Connection test successful!");
                    UpdateStatus("Connection test successful");
                }
            }
            catch (Exception ex)
            {
                LogMessage($"❌ Connection test failed: {ex.Message}");
                UpdateStatus("Connection test failed");
            }
            finally
            {
                btnTestConnection.Enabled = true;
            }
        }

        private async void BtnExport_Click(object sender, EventArgs e)
        {
            if (cmbConnectionString.SelectedItem == null)
            {
                LogMessage("Please select a connection string first.");
                return;
            }

            if (string.IsNullOrWhiteSpace(txtSqlQuery.Text))
            {
                LogMessage("Please enter a SQL query.");
                return;
            }

            if (string.IsNullOrWhiteSpace(txtExportPath.Text))
            {
                LogMessage("Please select an export folder.");
                return;
            }

            if (!Directory.Exists(txtExportPath.Text))
            {
                LogMessage("Selected export folder does not exist.");
                return;
            }

            btnExport.Enabled = false;
            progressBar.Value = 0;
            UpdateStatus("Starting export...");

            try
            {
                var connectionString = cmbConnectionString.SelectedItem.ToString().Split(" - ")[1];
                var sqlQuery = txtSqlQuery.Text.Trim();
                var maxRowsPerSheet = (int)numMaxRowsPerSheet.Value;

                LogMessage($"Starting export with max {maxRowsPerSheet:N0} rows per sheet...");

                using (var connection = new OracleConnection(connectionString))
                {
                    await connection.OpenAsync();
                    LogMessage("Connected to database successfully.");

                    // First, get total row count
                    var countQuery = $"SELECT COUNT(*) FROM ({sqlQuery})";
                    using (var countCmd = new OracleCommand(countQuery, connection))
                    {
                        var totalRows = Convert.ToInt64(await countCmd.ExecuteScalarAsync());
                        LogMessage($"Total rows to export: {totalRows:N0}");

                        if (totalRows == 0)
                        {
                            LogMessage("No data found to export.");
                            UpdateStatus("No data found");
                            return;
                        }

                        // Calculate number of sheets needed
                        var numberOfSheets = (int)Math.Ceiling((double)totalRows / maxRowsPerSheet);
                        LogMessage($"Will create {numberOfSheets} sheet(s)");

                        // Create Excel file
                        var fileName = $"Export_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";
                        var filePath = Path.Combine(txtExportPath.Text, fileName);

                        using (var package = new ExcelPackage())
                        {
                            for (int sheetIndex = 0; sheetIndex < numberOfSheets; sheetIndex++)
                            {
                                var sheetName = numberOfSheets == 1 ? "Data" : $"Data_{sheetIndex + 1}";
                                var worksheet = package.Workbook.Worksheets.Add(sheetName);
                                
                                LogMessage($"Processing sheet {sheetIndex + 1}/{numberOfSheets}: {sheetName}");

                                // Calculate offset and fetch size for this sheet
                                var offset = sheetIndex * maxRowsPerSheet + 1;
                                var fetchSize = Math.Min(maxRowsPerSheet, (int)(totalRows - (sheetIndex * maxRowsPerSheet)));

                                var pagedQuery = $@"
                                    SELECT * FROM (
                                        SELECT ROWNUM as row_num, t.* FROM ({sqlQuery}) t
                                        WHERE ROWNUM <= {offset + fetchSize - 1}
                                    ) WHERE row_num >= {offset}";

                                using (var cmd = new OracleCommand(pagedQuery, connection))
                                {
                                    using (var reader = await cmd.ExecuteReaderAsync())
                                    {
                                        // Write headers
                                        for (int col = 0; col < reader.FieldCount; col++)
                                        {
                                            worksheet.Cells[1, col + 1].Value = reader.GetName(col);
                                            worksheet.Cells[1, col + 1].Style.Font.Bold = true;
                                        }

                                        // Write data
                                        int row = 2;
                                        int rowsInThisSheet = 0;
                                        
                                        while (await reader.ReadAsync() && rowsInThisSheet < maxRowsPerSheet)
                                        {
                                            for (int col = 0; col < reader.FieldCount; col++)
                                            {
                                                var value = reader.IsDBNull(col) ? "" : reader.GetValue(col).ToString();
                                                worksheet.Cells[row, col + 1].Value = value;
                                            }
                                            row++;
                                            rowsInThisSheet++;
                                            
                                            if (rowsInThisSheet % 10000 == 0)
                                            {
                                                var progress = (int)((double)(sheetIndex * maxRowsPerSheet + rowsInThisSheet) / totalRows * 100);
                                                UpdateProgressBar(progress);
                                                LogMessage($"Exported {rowsInThisSheet:N0} rows in current sheet...");
                                            }
                                        }
                                        
                                        LogMessage($"Completed sheet {sheetIndex + 1}: {rowsInThisSheet:N0} rows");
                                    }
                                }
                            }

                            // Save the file
                            await Task.Run(() => package.SaveAs(new FileInfo(filePath)));
                        }

                        LogMessage($"✅ Export completed successfully!");
                        LogMessage($"File saved: {filePath}");
                        UpdateStatus($"Export completed: {fileName}");
                        UpdateProgressBar(100);

                        // Open the file location
                        System.Diagnostics.Process.Start("explorer.exe", "/select," + filePath);
                    }
                }
            }
            catch (Exception ex)
            {
                LogMessage($"❌ Export failed: {ex.Message}");
                UpdateStatus("Export failed");
            }
            finally
            {
                btnExport.Enabled = true;
            }
        }

        private void BtnSelectExportFolder_Click(object sender, EventArgs e)
        {
            try
            {
                using (var folderDialog = new FolderBrowserDialog())
                {
                    folderDialog.Description = "Select folder to save Excel file";
                    folderDialog.SelectedPath = txtExportPath.Text;
                    folderDialog.ShowNewFolderButton = true;

                    if (folderDialog.ShowDialog() == DialogResult.OK)
                    {
                        txtExportPath.Text = folderDialog.SelectedPath;
                        LogMessage($"Export folder selected: {folderDialog.SelectedPath}");
                    }
                }
            }
            catch (Exception ex)
            {
                LogMessage($"Error selecting folder: {ex.Message}");
            }
        }

        private void BtnOpenLogs_Click(object sender, EventArgs e)
        {
            try
            {
                var logDirectory = Path.Combine(Application.StartupPath, "logs");
                if (Directory.Exists(logDirectory))
                {
                    System.Diagnostics.Process.Start("explorer.exe", logDirectory);
                }
                else
                {
                    LogMessage("Log directory not found.");
                }
            }
            catch (Exception ex)
            {
                LogMessage($"Error opening logs: {ex.Message}");
            }
        }

        private void LogMessage(string message)
        {
            LogMessageRequested?.Invoke(this, message);
            if (txtLog.InvokeRequired)
            {
                txtLog.Invoke(new Action(() => txtLog.AppendText($"{DateTime.Now:HH:mm:ss} - {message}{Environment.NewLine}")));
            }
            else
            {
                txtLog.AppendText($"{DateTime.Now:HH:mm:ss} - {message}{Environment.NewLine}");
            }
        }

        private void UpdateStatus(string status)
        {
            StatusUpdateRequested?.Invoke(this, status);
            if (lblStatus.InvokeRequired)
            {
                lblStatus.Invoke(new Action(() => lblStatus.Text = $"Status: {status}"));
            }
            else
            {
                lblStatus.Text = $"Status: {status}";
            }
        }

        private void UpdateProgressBar(int value)
        {
            if (progressBar.InvokeRequired)
            {
                progressBar.Invoke(new Action(() => progressBar.Value = Math.Min(100, Math.Max(0, value))));
            }
            else
            {
                progressBar.Value = Math.Min(100, Math.Max(0, value));
            }
        }
    }
}
