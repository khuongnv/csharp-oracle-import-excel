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
    public partial class ImportTab : UserControl
    {
        // Events để giao tiếp với MainForm
        public event EventHandler<string> LogMessageRequested;
        public event EventHandler<string> StatusUpdateRequested;
        public event EventHandler<AppConfig> ConfigurationSaveRequested;
        public event EventHandler<AppConfig> ConfigurationLoadRequested;

        // Controls
        private ComboBox cmbConnectionString;
        private TextBox txtExcelFilePath;
        private TextBox txtTableName;
        private Button btnSelectExcelFile;
        private Button btnImport;
        private Button btnTestConnection;
        private Button btnPreviewData;
        private ProgressBar progressBar;
        private Label lblStatus;
        private Label lblLog;
        private RichTextBox txtLog;
        private NumericUpDown numBatchSize;
        private CheckBox chkHasHeader;
        private ComboBox cmbSheetSelection;
        private Button btnOpenLogs;

        private AppConfig config;

        public ImportTab()
        {
            InitializeComponent();
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            LoadConfiguration();
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();

            // Connection String
            var lblConnectionString = new Label
            {
                Text = "Oracle Connection String:",
                Location = new System.Drawing.Point(20, 20),
                Size = new System.Drawing.Size(200, 20)
            };
            this.Controls.Add(lblConnectionString);

            cmbConnectionString = new ComboBox
            {
                Location = new System.Drawing.Point(20, 45),
                Size = new System.Drawing.Size(600, 25),
                DropDownStyle = ComboBoxStyle.DropDownList,
                Text = "Data Source=localhost:1521/XE;User Id=username;Password=password;"
            };
            cmbConnectionString.SelectedIndexChanged += CmbConnectionString_SelectedIndexChanged;
            this.Controls.Add(cmbConnectionString);

            // Test Connection Button
            btnTestConnection = new Button
            {
                Text = "Test Connection",
                Location = new System.Drawing.Point(640, 43),
                Size = new System.Drawing.Size(120, 30)
            };
            btnTestConnection.Click += BtnTestConnection_Click;
            this.Controls.Add(btnTestConnection);

            // Excel File Selection
            var lblExcelFile = new Label
            {
                Text = "Excel File:",
                Location = new System.Drawing.Point(20, 80),
                Size = new System.Drawing.Size(100, 20)
            };
            this.Controls.Add(lblExcelFile);

            txtExcelFilePath = new TextBox
            {
                Location = new System.Drawing.Point(130, 78),
                Size = new System.Drawing.Size(500, 20),
                ReadOnly = true
            };
            this.Controls.Add(txtExcelFilePath);

            btnSelectExcelFile = new Button
            {
                Text = "Select File",
                Location = new System.Drawing.Point(640, 76),
                Size = new System.Drawing.Size(120, 25)
            };
            btnSelectExcelFile.Click += BtnSelectExcelFile_Click;
            this.Controls.Add(btnSelectExcelFile);

            // Table Name
            var lblTableName = new Label
            {
                Text = "Table Name:",
                Location = new System.Drawing.Point(20, 110),
                Size = new System.Drawing.Size(100, 20)
            };
            this.Controls.Add(lblTableName);

            txtTableName = new TextBox
            {
                Location = new System.Drawing.Point(130, 108),
                Size = new System.Drawing.Size(200, 20),
                Text = "EXCEL_IMPORT"
            };
            this.Controls.Add(txtTableName);

            // Sheet Selection
            var lblSheetSelection = new Label
            {
                Text = "Sheet Selection:",
                Location = new System.Drawing.Point(350, 110),
                Size = new System.Drawing.Size(100, 20)
            };
            this.Controls.Add(lblSheetSelection);

            cmbSheetSelection = new ComboBox
            {
                Location = new System.Drawing.Point(460, 108),
                Size = new System.Drawing.Size(200, 20),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            cmbSheetSelection.SelectedIndexChanged += CmbSheetSelection_SelectedIndexChanged;
            this.Controls.Add(cmbSheetSelection);

            // Batch Size
            var lblBatchSize = new Label
            {
                Text = "Batch Size:",
                Location = new System.Drawing.Point(20, 140),
                Size = new System.Drawing.Size(100, 20)
            };
            this.Controls.Add(lblBatchSize);

            numBatchSize = new NumericUpDown
            {
                Location = new System.Drawing.Point(130, 138),
                Size = new System.Drawing.Size(80, 20),
                Minimum = 1,
                Maximum = 10000,
                Value = 100
            };
            this.Controls.Add(numBatchSize);

            // Has Header Checkbox
            chkHasHeader = new CheckBox
            {
                Text = "First row contains headers",
                Location = new System.Drawing.Point(230, 140),
                Size = new System.Drawing.Size(200, 20),
                Checked = true
            };
            this.Controls.Add(chkHasHeader);

            // Preview Button
            btnPreviewData = new Button
            {
                Text = "Preview Data",
                Location = new System.Drawing.Point(450, 135),
                Size = new System.Drawing.Size(100, 30)
            };
            btnPreviewData.Click += BtnPreviewData_Click;
            this.Controls.Add(btnPreviewData);

            // Import Button
            btnImport = new Button
            {
                Text = "Import to Oracle",
                Location = new System.Drawing.Point(560, 135),
                Size = new System.Drawing.Size(120, 30),
                BackColor = System.Drawing.Color.LightGreen
            };
            btnImport.Click += BtnImport_Click;
            this.Controls.Add(btnImport);

            // Progress Bar
            progressBar = new ProgressBar
            {
                Location = new System.Drawing.Point(20, 180),
                Size = new System.Drawing.Size(760, 20)
            };
            this.Controls.Add(progressBar);

            // Status Label
            lblStatus = new Label
            {
                Text = "Ready",
                Location = new System.Drawing.Point(20, 210),
                Size = new System.Drawing.Size(200, 20)
            };
            this.Controls.Add(lblStatus);

            // Open Logs Button
            btnOpenLogs = new Button
            {
                Text = "Open Logs",
                Location = new System.Drawing.Point(700, 210),
                Size = new System.Drawing.Size(80, 25)
            };
            btnOpenLogs.Click += BtnOpenLogs_Click;
            this.Controls.Add(btnOpenLogs);

            // Log Label
            lblLog = new Label
            {
                Text = "Log:",
                Location = new System.Drawing.Point(20, 240),
                Size = new System.Drawing.Size(100, 20)
            };
            this.Controls.Add(lblLog);

            // Log TextBox
            txtLog = new RichTextBox
            {
                Location = new System.Drawing.Point(20, 260),
                Size = new System.Drawing.Size(840, 270),
                ReadOnly = true,
                Font = new System.Drawing.Font("Consolas", 9)
            };
            this.Controls.Add(txtLog);

            this.ResumeLayout(false);
        }

        // Public methods để MainForm có thể gọi
        public void RefreshConnectionComboBox(List<ConnectionStringItem> connections)
        {
            cmbConnectionString.Items.Clear();
            foreach (var connection in connections.OrderBy(x => x.Order))
            {
                cmbConnectionString.Items.Add(new ConnectionStringItem
                {
                    Id = connection.Id,
                    Name = connection.Name,
                    ConnectionString = connection.ConnectionString,
                    Order = connection.Order
                });
            }

            // Select the saved connection
            if (!string.IsNullOrEmpty(config.SelectedConnectionId))
            {
                for (int i = 0; i < cmbConnectionString.Items.Count; i++)
                {
                    var item = cmbConnectionString.Items[i] as ConnectionStringItem;
                    if (item != null && item.Id == config.SelectedConnectionId)
                    {
                        cmbConnectionString.SelectedIndex = i;
                        break;
                    }
                }
            }
            else if (cmbConnectionString.Items.Count > 0)
            {
                cmbConnectionString.SelectedIndex = 0;
            }
        }

        public void SetConfig(AppConfig sharedConfig)
        {
            config = sharedConfig;
            
            txtTableName.Text = config.TableName;
            chkHasHeader.Checked = config.HasHeader;
            numBatchSize.Value = config.BatchSize;
            txtExcelFilePath.Text = config.LastExcelFilePath;
            
            // Enable preview button if excel file exists
            btnPreviewData.Enabled = !string.IsNullOrEmpty(config.LastExcelFilePath) && File.Exists(config.LastExcelFilePath);
            
            // Load connection strings
            RefreshConnectionComboBox(config.ConnectionStrings);
        }

        public void LoadConfiguration()
        {
            try
            {
                config = AppConfig.Load();
                
                txtTableName.Text = config.TableName;
                chkHasHeader.Checked = config.HasHeader;
                numBatchSize.Value = config.BatchSize;
                txtExcelFilePath.Text = config.LastExcelFilePath;
                
                // Enable preview button if excel file exists
                btnPreviewData.Enabled = !string.IsNullOrEmpty(config.LastExcelFilePath) && File.Exists(config.LastExcelFilePath);
                
                // Load sheet selection if excel file exists
                if (btnPreviewData.Enabled)
                {
                    LoadSheetSelection();
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
            var selectedItem = cmbConnectionString.SelectedItem as ConnectionStringItem;
            config.ConnectionString = selectedItem?.ConnectionString ?? cmbConnectionString.Text;
            config.SelectedConnectionId = selectedItem?.Id ?? "";
            config.TableName = txtTableName.Text;
            config.HasHeader = chkHasHeader.Checked;
            config.BatchSize = (int)numBatchSize.Value;
            config.LastExcelFilePath = txtExcelFilePath.Text;
            config.SelectedSheetIndex = cmbSheetSelection.SelectedIndex;
            
            config.Save();
            FileLogger.LogInfo("Configuration saved");
        }

        // Event Handlers
        private void CmbConnectionString_SelectedIndexChanged(object sender, EventArgs e)
        {
            SaveConfiguration();
        }

        private void CmbSheetSelection_SelectedIndexChanged(object sender, EventArgs e)
        {
            SaveConfiguration();
        }

        private void BtnTestConnection_Click(object sender, EventArgs e)
        {
            try
            {
                var selectedItem = cmbConnectionString.SelectedItem as ConnectionStringItem;
                var connectionString = selectedItem?.ConnectionString ?? cmbConnectionString.Text;
                
                using (var connection = new OracleConnection(connectionString))
                {
                    connection.Open();
                    LogMessage("✓ Kết nối Oracle thành công!");
                    FileLogger.LogSuccess($"Oracle connection test successful: {connectionString}");
                    UpdateStatus("Connection successful", System.Drawing.Color.Green);
                    
                    // Save configuration after successful connection test
                    SaveConfiguration();
                }
            }
            catch (Exception ex)
            {
                LogMessage($"✗ Lỗi kết nối: {ex.Message}");
                FileLogger.LogError($"Oracle connection test failed: {cmbConnectionString.Text}", ex);
                UpdateStatus("Connection failed", System.Drawing.Color.Red);
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
                    config.LastExcelFilePath = openFileDialog.FileName;
                    SaveConfiguration();
                    
                    // Enable preview button
                    btnPreviewData.Enabled = true;
                    
                    // Load sheet selection
                    LoadSheetSelection();
                    
                    LogMessage($"Đã chọn file Excel: {Path.GetFileName(openFileDialog.FileName)}");
                    FileLogger.LogInfo($"Excel file selected: {openFileDialog.FileName}");
                }
            }
        }

        private void BtnPreviewData_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtExcelFilePath.Text) || !File.Exists(txtExcelFilePath.Text))
            {
                MessageBox.Show("Vui lòng chọn file Excel trước!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                using (var package = new ExcelPackage(new FileInfo(txtExcelFilePath.Text)))
                {
                    var worksheetCount = package.Workbook.Worksheets.Count;
                    var previewText = new StringBuilder();
                    
                    if (cmbSheetSelection.SelectedIndex == cmbSheetSelection.Items.Count - 1) // "Tất cả sheets"
                    {
                        previewText.AppendLine($"=== PREVIEW TẤT CẢ {worksheetCount} SHEETS ===\n");
                        
                        for (int i = 0; i < worksheetCount; i++)
                        {
                            var worksheet = package.Workbook.Worksheets[i];
                            PreviewSingleSheet(previewText, worksheet, i, worksheetCount);
                        }
                    }
                    else
                    {
                        var selectedSheetIndex = cmbSheetSelection.SelectedIndex;
                        var worksheet = package.Workbook.Worksheets[selectedSheetIndex];
                        PreviewSingleSheet(previewText, worksheet, selectedSheetIndex, worksheetCount);
                    }
                    
                    // Hiển thị preview trong MessageBox
                    var previewForm = new Form
                    {
                        Text = "Excel Data Preview",
                        Size = new System.Drawing.Size(800, 600),
                        StartPosition = FormStartPosition.CenterParent
                    };
                    
                    var textBox = new RichTextBox
                    {
                        Dock = DockStyle.Fill,
                        ReadOnly = true,
                        Font = new System.Drawing.Font("Consolas", 9),
                        Text = previewText.ToString()
                    };
                    
                    previewForm.Controls.Add(textBox);
                    previewForm.ShowDialog();
                }
            }
            catch (Exception ex)
            {
                LogMessage($"Lỗi preview dữ liệu: {ex.Message}");
                FileLogger.LogError("Error previewing Excel data", ex);
                MessageBox.Show($"Lỗi preview dữ liệu: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnImport_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtExcelFilePath.Text) || !File.Exists(txtExcelFilePath.Text))
            {
                MessageBox.Show("Vui lòng chọn file Excel trước!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (string.IsNullOrEmpty(txtTableName.Text))
            {
                MessageBox.Show("Vui lòng nhập tên bảng!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Disable import button to prevent multiple clicks
            btnImport.Enabled = false;
            progressBar.Value = 0;
            progressBar.Visible = true;

            // Run import in background
            Task.Run(async () =>
            {
                try
                {
                    await ImportExcelToOracle();
                }
                catch (Exception ex)
                {
                    LogMessage($"Lỗi import: {ex.Message}");
                    FileLogger.LogError("Error during import", ex);
                }
                finally
                {
                    // Re-enable import button on UI thread
                    this.Invoke(new Action(() =>
                    {
                        btnImport.Enabled = true;
                        progressBar.Visible = false;
                        UpdateStatus("Import completed", System.Drawing.Color.Blue);
                    }));
                }
            });
        }

        private void BtnOpenLogs_Click(object sender, EventArgs e)
        {
            try
            {
                var logPath = FileLogger.GetLogFilePath();
                if (File.Exists(logPath))
                {
                    System.Diagnostics.Process.Start("notepad.exe", logPath);
                }
                else
                {
                    MessageBox.Show("Không tìm thấy file log!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                LogMessage($"Lỗi mở file log: {ex.Message}");
                FileLogger.LogError("Error opening log file", ex);
            }
        }

        // Helper methods
        private void LogMessage(string message)
        {
            if (InvokeRequired)
            {
                Invoke(new Action<string>(LogMessage), message);
                return;
            }

            var logEntry = $"[{DateTime.Now:HH:mm:ss}] {message}\n";
            txtLog.AppendText(logEntry);
            txtLog.ScrollToCaret();
            
            LogMessageRequested?.Invoke(this, message);
        }

        private void UpdateStatus(string status, System.Drawing.Color color)
        {
            if (InvokeRequired)
            {
                Invoke(new Action<string, System.Drawing.Color>(UpdateStatus), status, color);
                return;
            }

            lblStatus.Text = status;
            lblStatus.ForeColor = color;
            
            StatusUpdateRequested?.Invoke(this, status);
        }

        private void LoadSheetSelection()
        {
            if (string.IsNullOrWhiteSpace(txtExcelFilePath.Text) || !File.Exists(txtExcelFilePath.Text))
                return;

            try
            {
                using (var package = new ExcelPackage(new FileInfo(txtExcelFilePath.Text)))
                {
                    var worksheetCount = package.Workbook.Worksheets.Count;
                    
                    cmbSheetSelection.Items.Clear();
                    
                    // Thêm các sheet vào ComboBox
                    for (int i = 0; i < worksheetCount; i++)
                    {
                        var worksheet = package.Workbook.Worksheets[i];
                        cmbSheetSelection.Items.Add($"Sheet {i + 1}: {worksheet.Name}");
                    }
                    
                    // Thêm option "Tất cả sheets"
                    cmbSheetSelection.Items.Add("Tất cả sheets");
                    
                    // Set default selection (sheet đầu tiên hoặc từ config)
                    if (config.SelectedSheetIndex >= 0 && config.SelectedSheetIndex < cmbSheetSelection.Items.Count)
                    {
                        cmbSheetSelection.SelectedIndex = config.SelectedSheetIndex;
                    }
                    else
                    {
                        cmbSheetSelection.SelectedIndex = 0; // Default: sheet đầu tiên
                    }
                    
                    LogMessage($"Đã load {worksheetCount} sheet(s) từ file Excel");
                    FileLogger.LogInfo($"Loaded {worksheetCount} sheets from Excel file");
                }
            }
            catch (Exception ex)
            {
                LogMessage($"Lỗi load sheet selection: {ex.Message}");
                FileLogger.LogError("Error loading sheet selection", ex);
            }
        }

        private void PreviewSingleSheet(StringBuilder previewText, ExcelWorksheet worksheet, int sheetIndex, int totalSheets)
        {
            var tableName = GetTableNameForSheet(sheetIndex, totalSheets);
            previewText.AppendLine($"=== SHEET {sheetIndex + 1}: {worksheet.Name} (Table: {tableName}) ===");
            
            var rowCount = worksheet.Dimension?.End.Row ?? 0;
            var colCount = worksheet.Dimension?.End.Column ?? 0;
            
            if (rowCount == 0 || colCount == 0)
            {
                previewText.AppendLine("Sheet trống!");
                previewText.AppendLine();
                return;
            }
            
            var dataStartRow = chkHasHeader.Checked ? 2 : 1;
            var previewRows = Math.Min(10, rowCount - dataStartRow + 1);
            
            previewText.AppendLine($"Dữ liệu: {rowCount} dòng, {colCount} cột");
            previewText.AppendLine($"Hiển thị {previewRows} dòng đầu tiên:");
            previewText.AppendLine();
            
            // Header
            if (chkHasHeader.Checked)
            {
                previewText.Append("Row | ");
                for (int col = 1; col <= colCount; col++)
                {
                    var cellValue = worksheet.Cells[1, col].Value?.ToString() ?? "";
                    previewText.Append($"{cellValue.PadRight(15)} | ");
                }
                previewText.AppendLine();
                previewText.AppendLine(new string('-', (colCount + 1) * 20));
            }
            
            // Data rows
            for (int row = dataStartRow; row < dataStartRow + previewRows; row++)
            {
                previewText.Append($"{row.ToString().PadRight(4)} | ");
                for (int col = 1; col <= colCount; col++)
                {
                    var cellValue = worksheet.Cells[row, col].Value?.ToString() ?? "";
                    previewText.Append($"{cellValue.PadRight(15)} | ");
                }
                previewText.AppendLine();
            }
            
            if (rowCount - dataStartRow + 1 > previewRows)
            {
                previewText.AppendLine($"... và {rowCount - dataStartRow + 1 - previewRows} dòng khác");
            }
            
            previewText.AppendLine();
        }

        private string GetTableNameForSheet(int sheetIndex, int totalSheets)
        {
            var baseTableName = CleanTableName(txtTableName.Text);
            
            if (totalSheets == 1)
            {
                return baseTableName;
            }
            else
            {
                return $"{baseTableName}_Sheet{sheetIndex + 1}";
            }
        }

        private string CleanTableName(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
                return "EXCEL_IMPORT";
            
            // Chuyển đổi tiếng Việt sang không dấu
            var cleanName = ConvertVietnameseToUnsigned(name);
            
            // Trim spaces
            cleanName = cleanName.Trim();
            
            // Thay thế các ký tự không phải chữ/số bằng underscore
            cleanName = System.Text.RegularExpressions.Regex.Replace(cleanName, @"[^a-zA-Z0-9]", "_");
            
            // Loại bỏ underscore liên tiếp
            cleanName = System.Text.RegularExpressions.Regex.Replace(cleanName, @"_+", "_");
            
            // Loại bỏ underscore ở đầu và cuối
            cleanName = cleanName.Trim('_');
            
            // Đảm bảo bắt đầu bằng chữ cái
            if (cleanName.Length > 0 && char.IsDigit(cleanName[0]))
            {
                cleanName = "TBL_" + cleanName;
            }
            
            // Đảm bảo không rỗng
            if (string.IsNullOrWhiteSpace(cleanName))
            {
                cleanName = "EXCEL_IMPORT";
            }
            
            // Giới hạn độ dài tên bảng (Oracle limit)
            if (cleanName.Length > 30)
            {
                cleanName = cleanName.Substring(0, 30);
            }
            
            return cleanName.ToUpper();
        }

        private string ConvertVietnameseToUnsigned(string input)
        {
            if (string.IsNullOrEmpty(input))
                return string.Empty;

            // Bảng chuyển đổi tiếng Việt sang không dấu
            var vietnameseMap = new Dictionary<char, char>
            {
                {'à', 'a'}, {'á', 'a'}, {'ạ', 'a'}, {'ả', 'a'}, {'ã', 'a'}, {'â', 'a'}, {'ầ', 'a'}, {'ấ', 'a'}, {'ậ', 'a'}, {'ẩ', 'a'}, {'ẫ', 'a'}, {'ă', 'a'}, {'ằ', 'a'}, {'ắ', 'a'}, {'ặ', 'a'}, {'ẳ', 'a'}, {'ẵ', 'a'},
                {'è', 'e'}, {'é', 'e'}, {'ẹ', 'e'}, {'ẻ', 'e'}, {'ẽ', 'e'}, {'ê', 'e'}, {'ề', 'e'}, {'ế', 'e'}, {'ệ', 'e'}, {'ể', 'e'}, {'ễ', 'e'},
                {'ì', 'i'}, {'í', 'i'}, {'ị', 'i'}, {'ỉ', 'i'}, {'ĩ', 'i'},
                {'ò', 'o'}, {'ó', 'o'}, {'ọ', 'o'}, {'ỏ', 'o'}, {'õ', 'o'}, {'ô', 'o'}, {'ồ', 'o'}, {'ố', 'o'}, {'ộ', 'o'}, {'ổ', 'o'}, {'ỗ', 'o'}, {'ơ', 'o'}, {'ờ', 'o'}, {'ớ', 'o'}, {'ợ', 'o'}, {'ở', 'o'}, {'ỡ', 'o'},
                {'ù', 'u'}, {'ú', 'u'}, {'ụ', 'u'}, {'ủ', 'u'}, {'ũ', 'u'}, {'ư', 'u'}, {'ừ', 'u'}, {'ứ', 'u'}, {'ự', 'u'}, {'ử', 'u'}, {'ữ', 'u'},
                {'ỳ', 'y'}, {'ý', 'y'}, {'ỵ', 'y'}, {'ỷ', 'y'}, {'ỹ', 'y'},
                {'đ', 'd'},
                {'À', 'A'}, {'Á', 'A'}, {'Ạ', 'A'}, {'Ả', 'A'}, {'Ã', 'A'}, {'Â', 'A'}, {'Ầ', 'A'}, {'Ấ', 'A'}, {'Ậ', 'A'}, {'Ẩ', 'A'}, {'Ẫ', 'A'}, {'Ă', 'A'}, {'Ằ', 'A'}, {'Ắ', 'A'}, {'Ặ', 'A'}, {'Ẳ', 'A'}, {'Ẵ', 'A'},
                {'È', 'E'}, {'É', 'E'}, {'Ẹ', 'E'}, {'Ẻ', 'E'}, {'Ẽ', 'E'}, {'Ê', 'E'}, {'Ề', 'E'}, {'Ế', 'E'}, {'Ệ', 'E'}, {'Ể', 'E'}, {'Ễ', 'E'},
                {'Ì', 'I'}, {'Í', 'I'}, {'Ị', 'I'}, {'Ỉ', 'I'}, {'Ĩ', 'I'},
                {'Ò', 'O'}, {'Ó', 'O'}, {'Ọ', 'O'}, {'Ỏ', 'O'}, {'Õ', 'O'}, {'Ô', 'O'}, {'Ồ', 'O'}, {'Ố', 'O'}, {'Ộ', 'O'}, {'Ổ', 'O'}, {'Ỗ', 'O'}, {'Ơ', 'O'}, {'Ờ', 'O'}, {'Ớ', 'O'}, {'Ợ', 'O'}, {'Ở', 'O'}, {'Ỡ', 'O'},
                {'Ù', 'U'}, {'Ú', 'U'}, {'Ụ', 'U'}, {'Ủ', 'U'}, {'Ũ', 'U'}, {'Ư', 'U'}, {'Ừ', 'U'}, {'Ứ', 'U'}, {'Ự', 'U'}, {'Ử', 'U'}, {'Ữ', 'U'},
                {'Ỳ', 'Y'}, {'Ý', 'Y'}, {'Ỵ', 'Y'}, {'Ỷ', 'Y'}, {'Ỹ', 'Y'},
                {'Đ', 'D'}
            };

            var result = new StringBuilder();
            foreach (char c in input)
            {
                if (vietnameseMap.ContainsKey(c))
                {
                    result.Append(vietnameseMap[c]);
                }
                else
                {
                    result.Append(c);
                }
            }

            return result.ToString();
        }

        // Import methods - simplified for now
        private async Task ImportExcelToOracle()
        {
            LogMessage("Bắt đầu quá trình import...");
            FileLogger.LogInfo("Starting Excel import process");
            
            try
            {
                using (var package = new ExcelPackage(new FileInfo(txtExcelFilePath.Text)))
                {
                    var worksheetCount = package.Workbook.Worksheets.Count;
                    LogMessage($"Tìm thấy {worksheetCount} sheet(s) trong file Excel");
                    
                    if (cmbSheetSelection.SelectedIndex == cmbSheetSelection.Items.Count - 1) // "Tất cả sheets"
                    {
                        LogMessage("Đang import tất cả sheets...");
                        for (int i = 0; i < worksheetCount; i++)
                        {
                            var worksheet = package.Workbook.Worksheets[i];
                            await ProcessSingleSheet(worksheet, i, worksheetCount);
                        }
                    }
                    else
                    {
                        var selectedSheetIndex = cmbSheetSelection.SelectedIndex;
                        var worksheet = package.Workbook.Worksheets[selectedSheetIndex];
                        LogMessage($"Đang import sheet {selectedSheetIndex + 1}: {worksheet.Name}");
                        await ProcessSingleSheet(worksheet, selectedSheetIndex, 1);
                    }
                }
                
                LogMessage("✓ Import hoàn thành!");
                FileLogger.LogSuccess("Excel import completed successfully");
            }
            catch (Exception ex)
            {
                LogMessage($"✗ Lỗi import: {ex.Message}");
                FileLogger.LogError("Excel import failed", ex);
                throw;
            }
        }

        private async Task ProcessSingleSheet(ExcelWorksheet worksheet, int sheetIndex, int totalSheets)
        {
            LogMessage($"Đang xử lý sheet {sheetIndex + 1}: {worksheet.Name}");
            FileLogger.LogInfo($"Processing sheet {sheetIndex + 1}: {worksheet.Name}");

            try
            {
                var rowCount = worksheet.Dimension?.End.Row ?? 0;
                var colCount = worksheet.Dimension?.End.Column ?? 0;

                if (rowCount == 0 || colCount == 0)
                {
                    LogMessage($"Sheet {sheetIndex + 1} trống, bỏ qua");
                    return;
                }

                LogMessage($"Sheet {sheetIndex + 1}: {rowCount} dòng, {colCount} cột");

                // Tạo tên bảng dựa trên số lượng sheet
                string tableName = GetTableNameForSheet(sheetIndex, totalSheets);
                LogMessage($"Tên bảng cho sheet này: {tableName}");
                FileLogger.LogInfo($"Table name for sheet {sheetIndex + 1}: {tableName}");

                // Tạo bảng trong Oracle
                await CreateOracleTable(worksheet, tableName);

                // Insert dữ liệu với batch processing
                await InsertDataToOracle(worksheet, rowCount, colCount, tableName);
            }
            catch (Exception ex)
            {
                LogMessage($"Lỗi xử lý sheet {sheetIndex + 1}: {ex.Message}");
                FileLogger.LogError($"Error processing sheet {sheetIndex + 1}", ex);
                throw;
            }
        }

        private async Task CreateOracleTable(ExcelWorksheet worksheet, string tableName)
        {
            LogMessage($"Kiểm tra và tạo bảng '{tableName}' trong Oracle database...");
            FileLogger.LogInfo($"Creating Oracle table: {tableName}");

            var selectedItem = cmbConnectionString.SelectedItem as ConnectionStringItem;
            var connectionString = selectedItem?.ConnectionString ?? cmbConnectionString.Text;
            using (var connection = new OracleConnection(connectionString))
            {
                await connection.OpenAsync();

                // Check table tồn tại và có dữ liệu không
                var tableExists = await CheckTableExists(connection, tableName);
                if (tableExists)
                {
                    var hasData = await CheckTableHasData(connection, tableName);
                    if (hasData)
                    {
                        var result = MessageBox.Show($"Bảng '{tableName}' đã tồn tại và có dữ liệu. Bạn có muốn xóa và tạo mới không?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (result == DialogResult.Yes)
                        {
                            LogMessage($"Bảng '{tableName}' đã tồn tại và có dữ liệu - đang xóa và tạo mới");
                            await DropTable(connection, tableName);
                        }
                        else
                        {
                            LogMessage("Import bị hủy bởi người dùng");
                            return;
                        }
                    }
                    else
                    {
                        LogMessage($"Bảng '{tableName}' đã tồn tại nhưng trống - đang xóa và tạo mới");
                        await DropTable(connection, tableName);
                    }
                }

                // Tạo bảng mới
                var columnNames = GetColumnNames(worksheet);
                var createColumns = string.Join(", ", columnNames.Select(col => $"\"{col}\" VARCHAR2(4000)"));
                var createSql = $"CREATE TABLE \"{tableName}\" ({createColumns})";

                LogMessage($"SQL tạo bảng: {createSql}");
                FileLogger.LogInfo($"Create table SQL: {createSql}");

                try
                {
                    using (var createCommand = new OracleCommand(createSql, connection))
                    {
                        await createCommand.ExecuteNonQueryAsync();
                    }
                }
                catch (Exception ex)
                {
                    LogMessage($"Lỗi tạo bảng: {ex.Message}");
                    FileLogger.LogError($"Error creating table: {tableName}", ex);
                    throw;
                }

                // Thêm comment cho các cột
                await AddColumnComments(connection, tableName, columnNames, worksheet, chkHasHeader.Checked);

                LogMessage($"✓ Đã tạo bảng: {tableName}");
                FileLogger.LogSuccess($"Created Oracle table: {tableName} with columns: {string.Join(", ", columnNames)}");
            }
        }

        private async Task InsertDataToOracle(ExcelWorksheet worksheet, int rowCount, int colCount, string tableName)
        {
            LogMessage("Bắt đầu insert dữ liệu...");
            FileLogger.LogInfo("Starting data insertion");

            var selectedItem = cmbConnectionString.SelectedItem as ConnectionStringItem;
            var connectionString = selectedItem?.ConnectionString ?? cmbConnectionString.Text;
            using (var connection = new OracleConnection(connectionString))
            {
                await connection.OpenAsync();

                var columnNames = GetColumnNames(worksheet);
                var dataStartRow = chkHasHeader.Checked ? 2 : 1;
                var batchSize = (int)numBatchSize.Value;

                // Tạo SQL insert
                var insertColumns = string.Join(", ", columnNames.Select(col => $"\"{col}\""));
                var insertValues = string.Join(", ", columnNames.Select((col, index) => $":col{index}"));
                var insertSql = $"INSERT INTO \"{tableName}\" ({insertColumns}) VALUES ({insertValues})";

                LogMessage($"SQL insert: {insertSql}");
                FileLogger.LogInfo($"Insert SQL: {insertSql}");

                var totalRows = rowCount - dataStartRow + 1;
                var processedRows = 0;

                for (int startRow = dataStartRow; startRow <= rowCount; startRow += batchSize)
                {
                    var endRow = Math.Min(startRow + batchSize - 1, rowCount);
                    var currentBatchSize = endRow - startRow + 1;

                    try
                    {
                        using (var command = new OracleCommand(insertSql, connection))
                        {
                            // Tạo parameters cho batch
                            var parameters = new List<OracleParameter[]>();
                            for (int row = startRow; row <= endRow; row++)
                            {
                                var rowParams = new OracleParameter[colCount];
                                for (int col = 1; col <= colCount; col++)
                                {
                                    var cellValue = worksheet.Cells[row, col].Value?.ToString() ?? "";
                                    rowParams[col - 1] = new OracleParameter($":col{col - 1}", OracleDbType.Varchar2) { Value = cellValue };
                                }
                                parameters.Add(rowParams);
                            }

                            // Execute batch
                            command.ArrayBindCount = currentBatchSize;
                            for (int col = 0; col < colCount; col++)
                            {
                                var values = parameters.Select(p => p[col].Value).ToArray();
                                command.Parameters.Add(new OracleParameter($":col{col}", OracleDbType.Varchar2) { Value = values });
                            }

                            await command.ExecuteNonQueryAsync();
                        }

                        processedRows += currentBatchSize;
                        var progress = (int)((double)processedRows / totalRows * 100);
                        progressBar.Value = Math.Min(progress, 100);

                        LogMessage($"Đã insert {processedRows}/{totalRows} dòng ({progress}%)");
                    }
                    catch (Exception ex)
                    {
                        LogMessage($"Lỗi insert batch từ dòng {startRow} đến {endRow}: {ex.Message}");
                        FileLogger.LogError($"Error inserting batch from row {startRow} to {endRow}", ex);
                        throw;
                    }
                }

                LogMessage($"✓ Đã insert thành công {processedRows} dòng vào bảng '{tableName}'");
                FileLogger.LogSuccess($"Successfully inserted {processedRows} rows into table '{tableName}'");
            }
        }

        private List<string> GetColumnNames(ExcelWorksheet worksheet)
        {
            var columns = new List<string>();
            var colCount = worksheet.Dimension?.End.Column ?? 0;

            if (chkHasHeader.Checked)
            {
                // Lấy tên cột từ header
                for (int col = 1; col <= colCount; col++)
                {
                    var headerValue = worksheet.Cells[1, col].Value?.ToString() ?? "";
                    var cleanName = CleanColumnName(headerValue);
                    columns.Add(cleanName);
                }
            }
            else
            {
                // Tạo tên cột theo format A, B, C...
                for (int col = 1; col <= colCount; col++)
                {
                    if (col <= 26)
                    {
                        columns.Add(((char)('A' + col - 1)).ToString());
                    }
                    else
                    {
                        int first = (col - 1) / 26;
                        int second = (col - 1) % 26;
                        columns.Add($"{(char)('A' + first)}{(char)('A' + second)}");
                    }
                }
            }
            return columns;
        }

        private string CleanColumnName(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
                return "COLUMN_1";
            
            // Chuyển đổi tiếng Việt sang không dấu
            var cleanName = ConvertVietnameseToUnsigned(name);
            
            // Trim spaces
            cleanName = cleanName.Trim();
            
            // Thay thế các ký tự không phải chữ/số bằng underscore
            cleanName = System.Text.RegularExpressions.Regex.Replace(cleanName, @"[^a-zA-Z0-9]", "_");
            
            // Loại bỏ underscore liên tiếp
            cleanName = System.Text.RegularExpressions.Regex.Replace(cleanName, @"_+", "_");
            
            // Loại bỏ underscore ở đầu và cuối
            cleanName = cleanName.Trim('_');
            
            // Đảm bảo bắt đầu bằng chữ cái
            if (cleanName.Length > 0 && char.IsDigit(cleanName[0]))
            {
                cleanName = "COL_" + cleanName;
            }
            
            // Đảm bảo không rỗng
            if (string.IsNullOrWhiteSpace(cleanName))
            {
                cleanName = "COLUMN_1";
            }
            
            // Giới hạn độ dài tên cột (Oracle limit)
            if (cleanName.Length > 30)
            {
                cleanName = cleanName.Substring(0, 30);
            }
            
            return cleanName.ToUpper();
        }

        private async Task AddColumnComments(OracleConnection connection, string tableName, List<string> columnNames, ExcelWorksheet worksheet, bool hasHeader)
        {
            LogMessage("Đang thêm comment cho các cột...");
            FileLogger.LogInfo("Adding column comments");

            try
            {
                for (int col = 0; col < columnNames.Count; col++)
                {
                    string originalColumnName;
                    
                    if (hasHeader)
                    {
                        // Lấy tên gốc từ header Excel
                        originalColumnName = worksheet.Cells[1, col + 1].Value?.ToString() ?? "";
                        if (string.IsNullOrWhiteSpace(originalColumnName))
                        {
                            originalColumnName = $"Column {col + 1}";
                        }
                    }
                    else
                    {
                        // Tạo tên cột theo format A, B, C...
                        if (col < 26)
                        {
                            originalColumnName = ((char)('A' + col)).ToString();
                        }
                        else
                        {
                            int first = col / 26 - 1;
                            int second = col % 26;
                            originalColumnName = $"{(char)('A' + first)}{(char)('A' + second)}";
                        }
                    }

                    // Escape single quotes trong comment
                    var escapedComment = originalColumnName.Replace("'", "''");
                    
                    var commentSql = $"COMMENT ON COLUMN \"{tableName}\".\"{columnNames[col]}\" IS '{escapedComment}'";
                    
                    try
                    {
                        using (var commentCommand = new OracleCommand(commentSql, connection))
                        {
                            await commentCommand.ExecuteNonQueryAsync();
                        }
                    }
                    catch (Exception ex)
                    {
                        LogMessage($"⚠️ Không thể thêm comment cho cột {columnNames[col]}: {ex.Message}");
                        FileLogger.LogInfo($"WARNING: Failed to add comment for column {columnNames[col]}: {ex.Message}");
                    }
                }
                
                LogMessage("✓ Đã thêm comment cho các cột");
                FileLogger.LogSuccess("Column comments added successfully");
            }
            catch (Exception ex)
            {
                LogMessage($"⚠️ Lỗi thêm comment: {ex.Message}");
                FileLogger.LogError("Error adding column comments", ex);
            }
        }

        private async Task<bool> CheckTableExists(OracleConnection connection, string tableName)
        {
            var checkSql = @"
                SELECT COUNT(*) 
                FROM USER_TABLES 
                WHERE TABLE_NAME = UPPER(:tableName)";

            try
            {
                using (var command = new OracleCommand(checkSql, connection))
                {
                    command.Parameters.Add(":tableName", OracleDbType.Varchar2).Value = tableName;
                    var count = await command.ExecuteScalarAsync();
                    return Convert.ToInt32(count) > 0;
                }
            }
            catch (Exception ex)
            {
                LogMessage($"Lỗi kiểm tra bảng tồn tại: {ex.Message}");
                FileLogger.LogError("Error checking table existence", ex);
                return false;
            }
        }

        private async Task<bool> CheckTableHasData(OracleConnection connection, string tableName)
        {
            var countSql = $"SELECT COUNT(*) FROM \"{tableName}\"";
            
            try
            {
                using (var command = new OracleCommand(countSql, connection))
                {
                    var count = await command.ExecuteScalarAsync();
                    return Convert.ToInt32(count) > 0;
                }
            }
            catch (Exception ex)
            {
                LogMessage($"Lỗi kiểm tra dữ liệu bảng: {ex.Message}");
                FileLogger.LogError("Error checking table data", ex);
                return false;
            }
        }

        private async Task DropTable(OracleConnection connection, string tableName)
        {
            var dropSql = $"DROP TABLE \"{tableName}\"";
            
            try
            {
                using (var command = new OracleCommand(dropSql, connection))
                {
                    await command.ExecuteNonQueryAsync();
                }
                LogMessage($"Đã xóa bảng '{tableName}'");
                FileLogger.LogInfo($"Dropped table: {tableName}");
            }
            catch (Exception ex)
            {
                LogMessage($"Lỗi xóa bảng: {ex.Message}");
                FileLogger.LogError("Error dropping table", ex);
                throw;
            }
        }
    }
}
