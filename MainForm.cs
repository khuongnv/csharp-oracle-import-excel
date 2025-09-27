using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using Oracle.ManagedDataAccess.Client;
using OfficeOpenXml;

namespace ExcelToOracleImporter
{
    public partial class MainForm : Form
    {
        private TextBox txtConnectionString;
        private TextBox txtExcelFilePath;
        private TextBox txtTableName;
        private Button btnSelectExcelFile;
        private Button btnImport;
        private Button btnTestConnection;
        private Button btnPreviewData;
        private ProgressBar progressBar;
        private Label lblStatus;
        private RichTextBox txtLog;
        private NumericUpDown numBatchSize;
        private CheckBox chkHasHeader;
        private AppConfig config;
        private Button btnOpenLogs;
        private MenuStrip menuStrip;

        public MainForm()
        {
            InitializeComponent();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            
            // Load configuration
            config = AppConfig.Load();
            LoadConfiguration();
            
            // Log application start
            FileLogger.LogInfo("Excel to Oracle Importer started");
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();

            // Form properties
            this.Text = "Excel to Oracle Database Importer v2.1.0";
            this.Size = new System.Drawing.Size(900, 700);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.MaximizeBox = false;

            // Menu Strip
            menuStrip = new MenuStrip();
            var helpMenu = new ToolStripMenuItem("Help");
            var aboutMenuItem = new ToolStripMenuItem("About", null, (s, e) => ShowAboutDialog());
            helpMenu.DropDownItems.Add(aboutMenuItem);
            menuStrip.Items.Add(helpMenu);
            this.Controls.Add(menuStrip);
            this.MainMenuStrip = menuStrip;

            // Connection String
            var lblConnectionString = new Label
            {
                Text = "Oracle Connection String:",
                Location = new System.Drawing.Point(20, 20),
                Size = new System.Drawing.Size(200, 20)
            };
            this.Controls.Add(lblConnectionString);

            txtConnectionString = new TextBox
            {
                Location = new System.Drawing.Point(20, 45),
                Size = new System.Drawing.Size(600, 25),
                Text = "Data Source=localhost:1521/XE;User Id=username;Password=password;"
            };
            this.Controls.Add(txtConnectionString);

            // Test Connection Button
            btnTestConnection = new Button
            {
                Text = "Test Connection",
                Location = new System.Drawing.Point(640, 43),
                Size = new System.Drawing.Size(120, 30)
            };
            btnTestConnection.Click += BtnTestConnection_Click;
            this.Controls.Add(btnTestConnection);

            // Table Name
            var lblTableName = new Label
            {
                Text = "Table Name (sẽ được tạo mới):",
                Location = new System.Drawing.Point(20, 80),
                Size = new System.Drawing.Size(200, 20)
            };
            this.Controls.Add(lblTableName);

            txtTableName = new TextBox
            {
                Location = new System.Drawing.Point(20, 105),
                Size = new System.Drawing.Size(300, 25),
                Text = "EXCEL_IMPORT"
            };
            this.Controls.Add(txtTableName);

            // Excel File Path
            var lblExcelFile = new Label
            {
                Text = "Excel File Path:",
                Location = new System.Drawing.Point(20, 140),
                Size = new System.Drawing.Size(200, 20)
            };
            this.Controls.Add(lblExcelFile);

            txtExcelFilePath = new TextBox
            {
                Location = new System.Drawing.Point(20, 165),
                Size = new System.Drawing.Size(500, 25),
                ReadOnly = true
            };
            this.Controls.Add(txtExcelFilePath);

            btnSelectExcelFile = new Button
            {
                Text = "Select Excel File",
                Location = new System.Drawing.Point(540, 163),
                Size = new System.Drawing.Size(120, 30)
            };
            btnSelectExcelFile.Click += BtnSelectExcelFile_Click;
            this.Controls.Add(btnSelectExcelFile);

            // Preview Data Button
            btnPreviewData = new Button
            {
                Text = "Preview Data",
                Location = new System.Drawing.Point(670, 163),
                Size = new System.Drawing.Size(120, 30),
                Enabled = false
            };
            btnPreviewData.Click += BtnPreviewData_Click;
            this.Controls.Add(btnPreviewData);

            // Has Header Checkbox
            chkHasHeader = new CheckBox
            {
                Text = "File có Header (dòng đầu tiên là tên cột)",
                Location = new System.Drawing.Point(20, 200),
                Size = new System.Drawing.Size(300, 20),
                Checked = true
            };
            this.Controls.Add(chkHasHeader);

            // Batch Size Configuration
            var lblBatchSize = new Label
            {
                Text = "Batch Size:",
                Location = new System.Drawing.Point(350, 200),
                Size = new System.Drawing.Size(80, 20)
            };
            this.Controls.Add(lblBatchSize);

            numBatchSize = new NumericUpDown
            {
                Location = new System.Drawing.Point(430, 198),
                Size = new System.Drawing.Size(80, 25),
                Minimum = 1,
                Maximum = 10000,
                Value = 100
            };
            this.Controls.Add(numBatchSize);

            // Import Button
            btnImport = new Button
            {
                Text = "Import Data",
                Location = new System.Drawing.Point(20, 235),
                Size = new System.Drawing.Size(120, 40),
                BackColor = System.Drawing.Color.LightGreen
            };
            btnImport.Click += BtnImport_Click;
            this.Controls.Add(btnImport);

            // Progress Bar
            progressBar = new ProgressBar
            {
                Location = new System.Drawing.Point(160, 245),
                Size = new System.Drawing.Size(500, 20),
                Visible = false
            };
            this.Controls.Add(progressBar);

            // Status Label
            lblStatus = new Label
            {
                Text = "Ready to import",
                Location = new System.Drawing.Point(20, 285),
                Size = new System.Drawing.Size(600, 20)
            };
            this.Controls.Add(lblStatus);

            // Log TextBox
            var lblLog = new Label
            {
                Text = "Import Log:",
                Location = new System.Drawing.Point(20, 315),
                Size = new System.Drawing.Size(100, 20)
            };
            this.Controls.Add(lblLog);

            // Open Logs Button
            btnOpenLogs = new Button
            {
                Text = "Open Logs Folder",
                Location = new System.Drawing.Point(750, 312),
                Size = new System.Drawing.Size(120, 25)
            };
            btnOpenLogs.Click += BtnOpenLogs_Click;
            this.Controls.Add(btnOpenLogs);

            txtLog = new RichTextBox
            {
                Location = new System.Drawing.Point(20, 340),
                Size = new System.Drawing.Size(840, 300),
                ReadOnly = true,
                Font = new System.Drawing.Font("Consolas", 9)
            };
            this.Controls.Add(txtLog);

            this.ResumeLayout(false);
        }

        private void LoadConfiguration()
        {
            // Load saved configuration
            txtConnectionString.Text = config.ConnectionString;
            txtTableName.Text = config.TableName;
            chkHasHeader.Checked = config.HasHeader;
            numBatchSize.Value = config.BatchSize;
            txtExcelFilePath.Text = config.LastExcelFilePath;
            
            // Enable preview button if excel file exists
            btnPreviewData.Enabled = !string.IsNullOrEmpty(config.LastExcelFilePath) && File.Exists(config.LastExcelFilePath);
        }

        private void SaveConfiguration()
        {
            config.ConnectionString = txtConnectionString.Text;
            config.TableName = txtTableName.Text;
            config.HasHeader = chkHasHeader.Checked;
            config.BatchSize = (int)numBatchSize.Value;
            config.LastExcelFilePath = txtExcelFilePath.Text;
            
            config.Save();
            FileLogger.LogInfo("Configuration saved");
        }

        private void BtnTestConnection_Click(object sender, EventArgs e)
        {
            try
            {
                using (var connection = new OracleConnection(txtConnectionString.Text))
                {
                    connection.Open();
                    LogMessage("✓ Kết nối Oracle thành công!");
                    FileLogger.LogSuccess($"Oracle connection test successful: {txtConnectionString.Text}");
                    lblStatus.Text = "Connection successful";
                    lblStatus.ForeColor = System.Drawing.Color.Green;
                    
                    // Save configuration after successful connection test
                    SaveConfiguration();
                }
            }
            catch (Exception ex)
            {
                LogMessage($"✗ Lỗi kết nối: {ex.Message}");
                FileLogger.LogError($"Oracle connection test failed: {txtConnectionString.Text}", ex);
                lblStatus.Text = "Connection failed";
                lblStatus.ForeColor = System.Drawing.Color.Red;
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
                    LogMessage($"✓ Đã chọn file: {Path.GetFileName(openFileDialog.FileName)}");
                    FileLogger.LogInfo($"Excel file selected: {openFileDialog.FileName}");
                    btnPreviewData.Enabled = true;
                    
                    // Save configuration when file is selected
                    SaveConfiguration();
                }
            }
        }

        private void BtnPreviewData_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtExcelFilePath.Text) || !File.Exists(txtExcelFilePath.Text))
            {
                MessageBox.Show("Vui lòng chọn file Excel hợp lệ trước!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                using (var package = new ExcelPackage(new FileInfo(txtExcelFilePath.Text)))
                {
                    var worksheet = package.Workbook.Worksheets[0];
                    var rowCount = worksheet.Dimension?.Rows ?? 0;
                    var colCount = worksheet.Dimension?.Columns ?? 0;

                    if (rowCount == 0 || colCount == 0)
                    {
                        MessageBox.Show("File Excel không có dữ liệu!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    var previewText = new StringBuilder();
                    previewText.AppendLine($"File: {Path.GetFileName(txtExcelFilePath.Text)}");
                    previewText.AppendLine($"Kích thước: {rowCount} dòng x {colCount} cột");
                    previewText.AppendLine();

                    // Preview first 10 rows
                    var maxRows = Math.Min(10, rowCount);
                    var startRow = chkHasHeader.Checked ? 1 : 1;

                    for (int row = startRow; row <= Math.Min(startRow + maxRows - 1, rowCount); row++)
                    {
                        var rowData = new List<string>();
                        for (int col = 1; col <= colCount; col++)
                        {
                            var cellValue = worksheet.Cells[row, col].Value?.ToString() ?? "";
                            rowData.Add(cellValue.Length > 50 ? cellValue.Substring(0, 50) + "..." : cellValue);
                        }
                        previewText.AppendLine($"Row {row}: {string.Join(" | ", rowData)}");
                    }

                    if (rowCount > maxRows)
                    {
                        previewText.AppendLine($"... và {rowCount - maxRows} dòng khác");
                    }

                    FileLogger.LogInfo($"Excel data preview requested for file: {txtExcelFilePath.Text}");
                    MessageBox.Show(previewText.ToString(), "Excel Data Preview", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                FileLogger.LogError($"Error reading Excel file for preview: {txtExcelFilePath.Text}", ex);
                MessageBox.Show($"Lỗi đọc file Excel: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async void BtnImport_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtConnectionString.Text))
            {
                MessageBox.Show("Vui lòng nhập Connection String!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (string.IsNullOrWhiteSpace(txtExcelFilePath.Text) || !File.Exists(txtExcelFilePath.Text))
            {
                MessageBox.Show("Vui lòng chọn file Excel hợp lệ!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (string.IsNullOrWhiteSpace(txtTableName.Text))
            {
                MessageBox.Show("Vui lòng nhập tên bảng!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Save configuration before import
            SaveConfiguration();

            btnImport.Enabled = false;
            progressBar.Visible = true;
            progressBar.Style = ProgressBarStyle.Marquee;
            lblStatus.Text = "Importing...";
            lblStatus.ForeColor = System.Drawing.Color.Blue;

            FileLogger.LogInfo($"Starting import process - Excel: {txtExcelFilePath.Text}, Table: {txtTableName.Text}, Connection: {txtConnectionString.Text}");

            try
            {
                await ImportExcelToOracle();
                LogMessage("✓ Import hoàn thành thành công!");
                FileLogger.LogSuccess($"Import completed successfully - Excel: {txtExcelFilePath.Text}, Table: {txtTableName.Text}");
                lblStatus.Text = "Import completed successfully";
                lblStatus.ForeColor = System.Drawing.Color.Green;
                MessageBox.Show("Import dữ liệu thành công!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                LogMessage($"✗ Lỗi import: {ex.Message}");
                FileLogger.LogError($"Import failed - Excel: {txtExcelFilePath.Text}, Table: {txtTableName.Text}", ex);
                lblStatus.Text = "Import failed";
                lblStatus.ForeColor = System.Drawing.Color.Red;
                MessageBox.Show($"Lỗi import: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                btnImport.Enabled = true;
                progressBar.Visible = false;
            }
        }

        private async Task ImportExcelToOracle()
        {
            LogMessage("Bắt đầu đọc file Excel...");
            FileLogger.LogInfo("Starting Excel file reading process");

            // Đọc file Excel
            using (var package = new ExcelPackage(new FileInfo(txtExcelFilePath.Text)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                var rowCount = worksheet.Dimension?.Rows ?? 0;
                var colCount = worksheet.Dimension?.Columns ?? 0;

                if (rowCount == 0 || colCount == 0)
                {
                    var errorMsg = "File Excel không có dữ liệu!";
                    FileLogger.LogError(errorMsg);
                    throw new Exception(errorMsg);
                }

                LogMessage($"Tìm thấy {rowCount} dòng, {colCount} cột");
                FileLogger.LogInfo($"Excel file contains {rowCount} rows and {colCount} columns");

                // Xác định tên cột
                List<string> columnNames;
                int dataStartRow = 1;

                if (chkHasHeader.Checked)
                {
                    // Đọc header từ dòng đầu tiên
                    columnNames = new List<string>();
                    for (int col = 1; col <= colCount; col++)
                    {
                        var headerValue = worksheet.Cells[1, col].Value?.ToString() ?? "";
                        if (string.IsNullOrWhiteSpace(headerValue))
                        {
                            columnNames.Add($"COLUMN_{col}");
                        }
                        else
                        {
                            // Làm sạch tên cột để phù hợp với Oracle
                            var cleanName = CleanColumnName(headerValue);
                            columnNames.Add(cleanName);
                        }
                    }
                    dataStartRow = 2;
                    LogMessage($"Sử dụng header từ Excel: {string.Join(", ", columnNames)}");
                    FileLogger.LogInfo($"Using Excel headers: {string.Join(", ", columnNames)}");
                }
                else
                {
                    // Tạo tên cột từ A đến Z
                    columnNames = GenerateColumnNames(colCount);
                    LogMessage($"Tạo tên cột tự động: {string.Join(", ", columnNames)}");
                    FileLogger.LogInfo($"Generated column names: {string.Join(", ", columnNames)}");
                }

                // Tạo bảng trong Oracle
                await CreateOracleTable(columnNames);

                // Insert dữ liệu với batch processing
                await InsertDataToOracle(worksheet, rowCount, colCount, columnNames, dataStartRow);
            }
        }

        private string CleanColumnName(string name)
        {
            // Loại bỏ ký tự đặc biệt và thay thế bằng underscore
            var cleanName = Regex.Replace(name, @"[^a-zA-Z0-9_]", "_");
            
            // Đảm bảo bắt đầu bằng chữ cái
            if (cleanName.Length > 0 && char.IsDigit(cleanName[0]))
            {
                cleanName = "COL_" + cleanName;
            }
            
            // Giới hạn độ dài tên cột
            if (cleanName.Length > 30)
            {
                cleanName = cleanName.Substring(0, 30);
            }
            
            return cleanName.ToUpper();
        }

        private List<string> GenerateColumnNames(int columnCount)
        {
            var columns = new List<string>();
            for (int i = 0; i < columnCount; i++)
            {
                if (i < 26)
                {
                    columns.Add(((char)('A' + i)).ToString());
                }
                else
                {
                    // Cho cột thứ 27 trở đi (AA, AB, AC...)
                    int first = i / 26 - 1;
                    int second = i % 26;
                    columns.Add($"{(char)('A' + first)}{(char)('A' + second)}");
                }
            }
            return columns;
        }

        private async Task CreateOracleTable(List<string> columnNames)
        {
            LogMessage("Kiểm tra và tạo bảng trong Oracle database...");
            FileLogger.LogInfo($"Creating Oracle table: {txtTableName.Text}");

            using (var connection = new OracleConnection(txtConnectionString.Text))
            {
                await connection.OpenAsync();

                // Check table tồn tại và có dữ liệu không
                var tableExists = await CheckTableExists(connection);
                if (tableExists)
                {
                    var hasData = await CheckTableHasData(connection);
                    
                    if (hasData)
                    {
                        LogMessage($"⚠️ Bảng '{txtTableName.Text}' đã tồn tại và có dữ liệu!");
                        FileLogger.LogInfo($"Table {txtTableName.Text} exists and contains data");
                        
                        var result = MessageBox.Show(
                            $"Bảng '{txtTableName.Text}' đã tồn tại và có dữ liệu.\n\n" +
                            "Bạn muốn:\n" +
                            "• YES: Xóa bảng cũ và tạo mới (MẤT DỮ LIỆU)\n" +
                            "• NO: Hủy import\n" +
                            "• RETRY: Nhập tên bảng khác",
                            "Bảng đã tồn tại",
                            MessageBoxButtons.YesNoCancel,
                            MessageBoxIcon.Warning);

                        if (result == DialogResult.Yes)
                        {
                            LogMessage($"Đang xóa bảng cũ: {txtTableName.Text}");
                            await DropTable(connection);
                        }
                        else if (result == DialogResult.Retry)
                        {
                            // User sẽ nhập tên bảng mới trong UI
                            throw new Exception("Vui lòng nhập tên bảng khác và thử lại!");
                        }
                        else
                        {
                            throw new Exception("Import đã bị hủy bởi user!");
                        }
                    }
                    else
                    {
                        LogMessage($"Bảng '{txtTableName.Text}' đã tồn tại nhưng trống - đang xóa và tạo mới");
                        await DropTable(connection);
                    }
                }

                // Tạo bảng mới
                var createColumns = string.Join(", ", columnNames.Select(col => $"{col} VARCHAR2(4000)"));
                var createSql = $"CREATE TABLE {txtTableName.Text} ({createColumns})";

                using (var createCommand = new OracleCommand(createSql, connection))
                {
                    await createCommand.ExecuteNonQueryAsync();
                }

                LogMessage($"✓ Đã tạo bảng: {txtTableName.Text}");
                FileLogger.LogSuccess($"Created Oracle table: {txtTableName.Text} with columns: {string.Join(", ", columnNames)}");
            }
        }

        private async Task<bool> CheckTableExists(OracleConnection connection)
        {
            var checkSql = $@"
                SELECT COUNT(*) 
                FROM USER_TABLES 
                WHERE TABLE_NAME = UPPER('{txtTableName.Text}')";

            using (var command = new OracleCommand(checkSql, connection))
            {
                var result = await command.ExecuteScalarAsync();
                return Convert.ToInt32(result) > 0;
            }
        }

        private async Task<bool> CheckTableHasData(OracleConnection connection)
        {
            var countSql = $"SELECT COUNT(*) FROM {txtTableName.Text}";

            try
            {
                using (var command = new OracleCommand(countSql, connection))
                {
                    var result = await command.ExecuteScalarAsync();
                    return Convert.ToInt32(result) > 0;
                }
            }
            catch
            {
                // Nếu không query được, giả sử có dữ liệu
                return true;
            }
        }

        private async Task DropTable(OracleConnection connection)
        {
            var dropSql = $"DROP TABLE {txtTableName.Text}";
            using (var dropCommand = new OracleCommand(dropSql, connection))
            {
                await dropCommand.ExecuteNonQueryAsync();
            }
            LogMessage($"✓ Đã xóa bảng cũ: {txtTableName.Text}");
            FileLogger.LogInfo($"Dropped existing table: {txtTableName.Text}");
        }

        private async Task InsertDataToOracle(ExcelWorksheet worksheet, int rowCount, int colCount, List<string> columnNames, int dataStartRow)
        {
            LogMessage("Bắt đầu insert dữ liệu theo batch...");
            FileLogger.LogInfo("Starting batch data insertion process");

            var batchSize = (int)numBatchSize.Value;
            var totalDataRows = rowCount - dataStartRow + 1;
            var processedRows = 0;

            // Auto-adjust batch size based on data volume for better performance
            if (totalDataRows > 10000 && batchSize < 500)
            {
                var suggestedBatchSize = Math.Min(1000, totalDataRows / 20);
                LogMessage($"💡 Gợi ý: File lớn ({totalDataRows} dòng) - nên dùng batch size {suggestedBatchSize} để tăng hiệu suất");
            }

            FileLogger.LogInfo($"Batch size: {batchSize}, Total data rows: {totalDataRows}");

            using (var connection = new OracleConnection(txtConnectionString.Text))
            {
                await connection.OpenAsync();

                var insertColumns = string.Join(", ", columnNames);
                var insertValues = string.Join(", ", columnNames.Select(col => $":{col}"));
                var insertSql = $"INSERT INTO {txtTableName.Text} ({insertColumns}) VALUES ({insertValues})";

                using (var command = new OracleCommand(insertSql, connection))
                {
                    // Chuẩn bị array parameters cho batch insert
                    var parameterArrays = new Dictionary<string, string[]>();
                    foreach (var col in columnNames)
                    {
                        parameterArrays[col] = new string[batchSize];
                        command.Parameters.Add($":{col}", OracleDbType.Varchar2, batchSize);
                    }

                    // Xử lý dữ liệu theo batch
                    for (int startRow = dataStartRow; startRow <= rowCount; startRow += batchSize)
                    {
                        var currentBatchSize = Math.Min(batchSize, rowCount - startRow + 1);
                        
                        // Chuẩn bị dữ liệu cho batch hiện tại
                        for (int i = 0; i < currentBatchSize; i++)
                        {
                            var currentRow = startRow + i;
                            for (int col = 1; col <= colCount; col++)
                            {
                                var cellValue = worksheet.Cells[currentRow, col].Value?.ToString() ?? "";
                                var cleanValue = cellValue.Length > 4000 ? cellValue.Substring(0, 4000) : cellValue;
                                parameterArrays[columnNames[col - 1]][i] = cleanValue;
                            }
                        }

                        // Set array values cho parameters
                        foreach (var col in columnNames)
                        {
                            command.Parameters[$":{col}"].Value = parameterArrays[col];
                        }

                        // Execute batch insert
                        command.ArrayBindCount = currentBatchSize;
                        await command.ExecuteNonQueryAsync();
                        
                        processedRows += currentBatchSize;

                        // Log progress
                        var percentage = (processedRows * 100) / totalDataRows;
                        LogMessage($"Đã import {processedRows}/{totalDataRows} dòng ({percentage}%) - Batch size: {currentBatchSize}");
                        FileLogger.LogInfo($"Import progress: {processedRows}/{totalDataRows} rows ({percentage}%) - Batch size: {currentBatchSize}");
                        
                        // Update progress bar
                        if (progressBar.InvokeRequired)
                        {
                            progressBar.Invoke(new Action(() => {
                                progressBar.Style = ProgressBarStyle.Continuous;
                                progressBar.Value = Math.Min(100, percentage);
                            }));
                        }
                        else
                        {
                            progressBar.Style = ProgressBarStyle.Continuous;
                            progressBar.Value = Math.Min(100, percentage);
                        }
                    }
                }

                LogMessage($"✓ Hoàn thành import {totalDataRows} dòng dữ liệu theo batch!");
                FileLogger.LogSuccess($"Batch data insertion completed successfully: {totalDataRows} rows imported to table {txtTableName.Text} using batch size {batchSize}");
            }
        }

        private void LogMessage(string message)
        {
            if (txtLog.InvokeRequired)
            {
                txtLog.Invoke(new Action<string>(LogMessage), message);
                return;
            }

            var logText = $"[{DateTime.Now:HH:mm:ss}] {message}\n";
            txtLog.AppendText(logText);
            txtLog.ScrollToCaret();
            
            // Also write to file
            FileLogger.Log(message);
        }

        private void BtnOpenLogs_Click(object sender, EventArgs e)
        {
            try
            {
                var logPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "logs");
                if (Directory.Exists(logPath))
                {
                    System.Diagnostics.Process.Start("explorer.exe", logPath);
                    FileLogger.LogInfo("Opened logs folder");
                }
                else
                {
                    MessageBox.Show("Logs folder not found!", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error opening logs folder: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                FileLogger.LogError("Error opening logs folder", ex);
            }
        }

        private void ShowAboutDialog()
        {
            try
            {
                using (var aboutForm = new AboutForm())
                {
                    aboutForm.ShowDialog(this);
                }
                FileLogger.LogInfo("About dialog opened");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error showing about dialog: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                FileLogger.LogError("Error showing about dialog", ex);
            }
        }
    }
}
