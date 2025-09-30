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
using OfficeOpenXml;
using Oracle.ManagedDataAccess.Client;

namespace ExcelToOracleImporter
{
    public partial class QuickImportTab : UserControl
    {
        // Events để giao tiếp với MainForm
        public event EventHandler<string> LogMessageRequested;
        public event EventHandler<string> StatusUpdateRequested;

        // Controls
        private TextBox txtConnectionString;
        private TextBox txtExcelFilePath;
        private TextBox txtTableName;
        private Button btnSelectExcelFile;
        private Button btnImport;
        private Button btnTestConnection;
        private Button btnPreviewData;
        private ProgressBar progressBar;
        private Label lblStatus;
        private Label lblLog;
        private TextBox txtLog;
        private NumericUpDown numBatchSize;
        private CheckBox chkHasHeader;
        private ComboBox cmbSheetSelection;
        private Button btnOpenLogs;

        private AppConfig config;

        public QuickImportTab()
        {
            InitializeComponent();
            LoadConfiguration();
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();

            // Connection String
            var lblConnectionString = new Label
            {
                Text = "Connection String:",
                Location = new System.Drawing.Point(20, 20),
                Size = new System.Drawing.Size(120, 20)
            };
            this.Controls.Add(lblConnectionString);

            txtConnectionString = new TextBox
            {
                Location = new System.Drawing.Point(150, 18),
                Size = new System.Drawing.Size(600, 20)
            };
            this.Controls.Add(txtConnectionString);

            btnTestConnection = new Button
            {
                Text = "Test Connection",
                Location = new System.Drawing.Point(760, 17),
                Size = new System.Drawing.Size(120, 25)
            };
            btnTestConnection.Click += BtnTestConnection_Click;
            this.Controls.Add(btnTestConnection);

            // Excel File Path
            var lblExcelFile = new Label
            {
                Text = "Excel File:",
                Location = new System.Drawing.Point(20, 60),
                Size = new System.Drawing.Size(120, 20)
            };
            this.Controls.Add(lblExcelFile);

            txtExcelFilePath = new TextBox
            {
                Location = new System.Drawing.Point(150, 58),
                Size = new System.Drawing.Size(600, 20),
                ReadOnly = true
            };
            this.Controls.Add(txtExcelFilePath);

            btnSelectExcelFile = new Button
            {
                Text = "Select File",
                Location = new System.Drawing.Point(760, 57),
                Size = new System.Drawing.Size(120, 25)
            };
            btnSelectExcelFile.Click += BtnSelectExcelFile_Click;
            this.Controls.Add(btnSelectExcelFile);

            // Table Name
            var lblTableName = new Label
            {
                Text = "Table Name:",
                Location = new System.Drawing.Point(20, 100),
                Size = new System.Drawing.Size(120, 20)
            };
            this.Controls.Add(lblTableName);

            txtTableName = new TextBox
            {
                Location = new System.Drawing.Point(150, 98),
                Size = new System.Drawing.Size(300, 20),
                Text = "EXCEL_IMPORT",
                Anchor = AnchorStyles.Top | AnchorStyles.Left
            };
            this.Controls.Add(txtTableName);

            // Has Header
            chkHasHeader = new CheckBox
            {
                Text = "Has Header Row",
                Location = new System.Drawing.Point(470, 98),
                Size = new System.Drawing.Size(120, 20),
                Checked = true
            };
            this.Controls.Add(chkHasHeader);

            // Batch Size
            var lblBatchSize = new Label
            {
                Text = "Batch Size:",
                Location = new System.Drawing.Point(20, 140),
                Size = new System.Drawing.Size(80, 20)
            };
            this.Controls.Add(lblBatchSize);

            numBatchSize = new NumericUpDown
            {
                Location = new System.Drawing.Point(150, 138),
                Size = new System.Drawing.Size(80, 20),
                Minimum = 1,
                Maximum = 10000,
                Value = 100
            };
            this.Controls.Add(numBatchSize);

            // Sheet Selection
            var lblSheet = new Label
            {
                Text = "Sheet:",
                Location = new System.Drawing.Point(20, 180),
                Size = new System.Drawing.Size(120, 20)
            };
            this.Controls.Add(lblSheet);

            cmbSheetSelection = new ComboBox
            {
                Location = new System.Drawing.Point(150, 178),
                Size = new System.Drawing.Size(300, 20),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            cmbSheetSelection.SelectedIndexChanged += CmbSheetSelection_SelectedIndexChanged;
            this.Controls.Add(cmbSheetSelection);

            // Buttons
            btnPreviewData = new Button
            {
                Text = "Preview Data",
                Location = new System.Drawing.Point(20, 220),
                Size = new System.Drawing.Size(100, 25)
            };
            btnPreviewData.Click += BtnPreviewData_Click;
            this.Controls.Add(btnPreviewData);

            btnImport = new Button
            {
                Text = "Import to Oracle",
                Location = new System.Drawing.Point(130, 220),
                Size = new System.Drawing.Size(120, 25),
                BackColor = Color.LightGreen
            };
            btnImport.Click += BtnImport_Click;
            this.Controls.Add(btnImport);

            btnOpenLogs = new Button
            {
                Text = "Open Logs",
                Location = new System.Drawing.Point(260, 220),
                Size = new System.Drawing.Size(100, 25)
            };
            btnOpenLogs.Click += BtnOpenLogs_Click;
            this.Controls.Add(btnOpenLogs);

            // Progress Bar
            progressBar = new ProgressBar
            {
                Location = new System.Drawing.Point(20, 260),
                Size = new System.Drawing.Size(960, 20),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            this.Controls.Add(progressBar);

            // Status Label
            lblStatus = new Label
            {
                Text = "Ready",
                Location = new System.Drawing.Point(20, 290),
                Size = new System.Drawing.Size(960, 20),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            this.Controls.Add(lblStatus);

            // Log Label
            lblLog = new Label
            {
                Text = "Log:",
                Location = new System.Drawing.Point(20, 320),
                Size = new System.Drawing.Size(100, 20),
                Anchor = AnchorStyles.Top | AnchorStyles.Left
            };
            this.Controls.Add(lblLog);

            // Log TextBox
            txtLog = new TextBox
            {
                Location = new System.Drawing.Point(20, 350),
                Size = new System.Drawing.Size(960, 200),
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                ReadOnly = true,
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
            };
            this.Controls.Add(txtLog);

            this.ResumeLayout(false);
        }

        // Public methods để MainForm có thể gọi
        public void SetConfig(AppConfig sharedConfig)
        {
            config = sharedConfig;
        }

        public void LoadConfiguration()
        {
            try
            {
                config = AppConfig.Load();
                if (config != null)
                {
                    txtTableName.Text = config.TableName;
                    numBatchSize.Value = config.BatchSize;
                    chkHasHeader.Checked = config.HasHeader;
                    txtConnectionString.Text = config.ConnectionString;
                }
            }
            catch (Exception ex)
            {
                LogMessage($"Lỗi tải cấu hình: {ex.Message}");
                FileLogger.LogError("Error loading configuration", ex);
            }
        }

        public void SaveConfiguration()
        {
            try
            {
                if (config != null)
                {
                    config.TableName = txtTableName.Text;
                    config.BatchSize = (int)numBatchSize.Value;
                    config.HasHeader = chkHasHeader.Checked;
                    config.ConnectionString = txtConnectionString.Text;
                    config.Save();
                    FileLogger.LogInfo("Configuration saved");
                }
            }
            catch (Exception ex)
            {
                LogMessage($"Lỗi lưu cấu hình: {ex.Message}");
                FileLogger.LogError("Error saving configuration", ex);
            }
        }

        // Event Handlers
        private void BtnTestConnection_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtConnectionString.Text))
            {
                MessageBox.Show("Vui lòng nhập chuỗi kết nối!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                using (var connection = new OracleConnection(txtConnectionString.Text))
                {
                    connection.Open();
                    MessageBox.Show("Kết nối thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    LogMessage($"✓ Test connection thành công");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi kết nối: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                LogMessage($"✗ Test connection thất bại: {ex.Message}");
            }
        }

        private void BtnSelectExcelFile_Click(object sender, EventArgs e)
        {
            using (var openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Excel Files|*.xlsx;*.xls";
                openFileDialog.Title = "Select Excel File";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    txtExcelFilePath.Text = openFileDialog.FileName;
                    LoadSheetNames();
                    LogMessage($"Đã chọn file: {Path.GetFileName(openFileDialog.FileName)}");
                }
            }
        }

        private void LoadSheetNames()
        {
            try
            {
                cmbSheetSelection.Items.Clear();
                using (var package = new ExcelPackage(new FileInfo(txtExcelFilePath.Text)))
                {
                    foreach (var worksheet in package.Workbook.Worksheets)
                    {
                        cmbSheetSelection.Items.Add(worksheet.Name);
                    }
                    if (cmbSheetSelection.Items.Count > 0)
                    {
                        cmbSheetSelection.SelectedIndex = 0;
                    }
                }
            }
            catch (Exception ex)
            {
                LogMessage($"Lỗi đọc file Excel: {ex.Message}");
            }
        }

        private void CmbSheetSelection_SelectedIndexChanged(object sender, EventArgs e)
        {
            LogMessage($"Đã chọn sheet: {cmbSheetSelection.SelectedItem}");
        }

        private void BtnPreviewData_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtExcelFilePath.Text))
            {
                MessageBox.Show("Vui lòng chọn file Excel!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                using (var package = new ExcelPackage(new FileInfo(txtExcelFilePath.Text)))
                {
                    var worksheet = package.Workbook.Worksheets[cmbSheetSelection.SelectedIndex];
                    var rowCount = worksheet.Dimension?.Rows ?? 0;
                    var colCount = worksheet.Dimension?.Columns ?? 0;

                    LogMessage($"Sheet '{worksheet.Name}': {rowCount} rows, {colCount} columns");

                    // Preview first 5 rows
                    var previewRows = Math.Min(5, rowCount);
                    for (int row = 1; row <= previewRows; row++)
                    {
                        var rowData = new List<string>();
                        for (int col = 1; col <= colCount; col++)
                        {
                            var cellValue = worksheet.Cells[row, col].Value?.ToString() ?? "";
                            rowData.Add(cellValue);
                        }
                        LogMessage($"Row {row}: {string.Join(" | ", rowData)}");
                    }
                }
            }
            catch (Exception ex)
            {
                LogMessage($"Lỗi preview data: {ex.Message}");
            }
        }

        private async void BtnImport_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtConnectionString.Text))
            {
                MessageBox.Show("Vui lòng nhập chuỗi kết nối!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (string.IsNullOrWhiteSpace(txtExcelFilePath.Text))
            {
                MessageBox.Show("Vui lòng chọn file Excel!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (string.IsNullOrWhiteSpace(txtTableName.Text))
            {
                MessageBox.Show("Vui lòng nhập tên bảng!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            btnImport.Enabled = false;
            progressBar.Value = 0;
            progressBar.Visible = true;

            try
            {
                await ImportDataAsync();
            }
            catch (Exception ex)
            {
                LogMessage($"Lỗi import: {ex.Message}");
                FileLogger.LogError("Import error", ex);
            }
            finally
            {
                btnImport.Enabled = true;
                progressBar.Visible = false;
            }
        }

        private async Task ImportDataAsync()
        {
            var connectionString = txtConnectionString.Text;
            var excelFilePath = txtExcelFilePath.Text;
            var tableName = txtTableName.Text;
            var batchSize = (int)numBatchSize.Value;
            var hasHeader = chkHasHeader.Checked;
            var sheetIndex = cmbSheetSelection.SelectedIndex;

            LogMessage("Bắt đầu import...");
            StatusUpdateRequested?.Invoke(this, "Importing...");

            try
            {
                using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
                {
                    var worksheet = package.Workbook.Worksheets[sheetIndex];
                    var totalRows = worksheet.Dimension?.Rows ?? 0;

                    if (totalRows == 0)
                    {
                        LogMessage("Không có dữ liệu trong sheet!");
                        return;
                    }

                    var startRow = hasHeader ? 2 : 1;
                    var dataRows = totalRows - startRow + 1;

                    LogMessage($"Tổng số dòng dữ liệu: {dataRows}");

                    using (var connection = new OracleConnection(connectionString))
                    {
                        connection.Open();
                        LogMessage("Đã kết nối đến Oracle database");

                        // Create table if not exists
                        await CreateTableIfNotExists(connection, worksheet, tableName, hasHeader);

                        // Import data in batches
                        var importedRows = 0;
                        for (int i = startRow; i <= totalRows; i += batchSize)
                        {
                            var batchEndRow = Math.Min(i + batchSize - 1, totalRows);
                            var batchData = ReadBatchData(worksheet, i, batchEndRow, hasHeader);
                            
                            if (batchData.Count > 0)
                            {
                                await InsertBatchData(connection, tableName, batchData);
                                importedRows += batchData.Count;
                            }

                            var progress = (int)((double)(i - startRow + 1) / dataRows * 100);
                            progressBar.Value = Math.Min(progress, 100);
                            LogMessage($"Đã import {importedRows}/{dataRows} dòng...");
                        }

                        LogMessage($"✓ Import hoàn thành! Tổng cộng: {importedRows} dòng");
                        StatusUpdateRequested?.Invoke(this, "Import completed");
                    }
                }
            }
            catch (Exception ex)
            {
                LogMessage($"✗ Lỗi import: {ex.Message}");
                throw;
            }
        }

        private async Task CreateTableIfNotExists(OracleConnection connection, ExcelWorksheet worksheet, string tableName, bool hasHeader)
        {
            var columnCount = worksheet.Dimension?.Columns ?? 0;
            var columnNames = new List<string>();

            if (hasHeader)
            {
                for (int col = 1; col <= columnCount; col++)
                {
                    var headerValue = worksheet.Cells[1, col].Value?.ToString() ?? $"Column{col}";
                    columnNames.Add($"\"{headerValue}\"");
                }
            }
            else
            {
                for (int col = 1; col <= columnCount; col++)
                {
                    columnNames.Add($"\"Column{col}\"");
                }
            }

            var createTableSql = $@"
                CREATE TABLE {tableName} (
                    ID NUMBER GENERATED BY DEFAULT AS IDENTITY PRIMARY KEY,
                    {string.Join(",\n                    ", columnNames.Select(name => $"{name} VARCHAR2(4000)"))}
                )";

            try
            {
                using (var command = new OracleCommand(createTableSql, connection))
                {
                    await command.ExecuteNonQueryAsync();
                    LogMessage($"Đã tạo bảng {tableName}");
                }
            }
            catch (OracleException ex) when (ex.Number == 955) // Table already exists
            {
                LogMessage($"Bảng {tableName} đã tồn tại");
            }
        }

        private List<Dictionary<string, object>> ReadBatchData(ExcelWorksheet worksheet, int startRow, int endRow, bool hasHeader)
        {
            var data = new List<Dictionary<string, object>>();
            var columnCount = worksheet.Dimension?.Columns ?? 0;

            for (int row = startRow; row <= endRow; row++)
            {
                var rowData = new Dictionary<string, object>();
                for (int col = 1; col <= columnCount; col++)
                {
                    var columnName = hasHeader ? 
                        worksheet.Cells[1, col].Value?.ToString() ?? $"Column{col}" : 
                        $"Column{col}";
                    
                    var cellValue = worksheet.Cells[row, col].Value;
                    rowData[columnName] = cellValue ?? DBNull.Value;
                }
                data.Add(rowData);
            }

            return data;
        }

        private async Task InsertBatchData(OracleConnection connection, string tableName, List<Dictionary<string, object>> batchData)
        {
            if (batchData.Count == 0) return;

            var firstRow = batchData[0];
            var columns = firstRow.Keys.ToList();
            var columnNames = string.Join(", ", columns.Select(c => $"\"{c}\""));
            var parameterNames = string.Join(", ", columns.Select(c => $":{c}"));

            var insertSql = $"INSERT INTO {tableName} ({columnNames}) VALUES ({parameterNames})";

            using (var command = new OracleCommand(insertSql, connection))
            {
                foreach (var row in batchData)
                {
                    command.Parameters.Clear();
                    foreach (var kvp in row)
                    {
                        command.Parameters.Add($":{kvp.Key}", kvp.Value);
                    }
                    await command.ExecuteNonQueryAsync();
                }
            }
        }

        private void BtnOpenLogs_Click(object sender, EventArgs e)
        {
            try
            {
                var logForm = new LogForm();
                logForm.Show();
            }
            catch (Exception ex)
            {
                LogMessage($"Lỗi mở log: {ex.Message}");
            }
        }

        private void LogMessage(string message)
        {
            var logText = $"[{DateTime.Now:HH:mm:ss}] {message}\r\n";
            txtLog.AppendText(logText);
            txtLog.SelectionStart = txtLog.Text.Length;
            txtLog.ScrollToCaret();
            
            LogMessageRequested?.Invoke(this, message);
        }
    }
}
