using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Oracle.ManagedDataAccess.Client;

namespace ExcelToOracleImporter
{
    public partial class ConnectionManagementTab : UserControl
    {
        // Events để giao tiếp với MainForm
        public event EventHandler<string> LogMessageRequested;
        public event EventHandler<AppConfig> ConfigurationSaveRequested;
        public event EventHandler<AppConfig> ConfigurationLoadRequested;

        // Controls
        private DataGridView dgvConnections;
        private TextBox txtConnectionName;
        private TextBox txtConnectionString;
        private NumericUpDown numConnectionOrder;
        private Button btnAddConnection;
        private Button btnUpdateConnection;
        private Button btnDeleteConnection;
        private Button btnTestSelectedConnection;

        private AppConfig config;

        public ConnectionManagementTab()
        {
            InitializeComponent();
            LoadConfiguration();
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();

            // DataGridView for connections
            dgvConnections = new DataGridView
            {
                Location = new System.Drawing.Point(20, 20),
                Size = new System.Drawing.Size(600, 300),
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                ReadOnly = true,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false
            };
            dgvConnections.SelectionChanged += DgvConnections_SelectionChanged;
            this.Controls.Add(dgvConnections);

            // Connection Name
            var lblConnectionName = new Label
            {
                Text = "Connection Name:",
                Location = new System.Drawing.Point(20, 340),
                Size = new System.Drawing.Size(100, 20)
            };
            this.Controls.Add(lblConnectionName);

            txtConnectionName = new TextBox
            {
                Location = new System.Drawing.Point(130, 338),
                Size = new System.Drawing.Size(200, 20)
            };
            this.Controls.Add(txtConnectionName);

            // Connection String
            var lblConnectionString = new Label
            {
                Text = "Connection String:",
                Location = new System.Drawing.Point(20, 370),
                Size = new System.Drawing.Size(100, 20)
            };
            this.Controls.Add(lblConnectionString);

            txtConnectionString = new TextBox
            {
                Location = new System.Drawing.Point(130, 368),
                Size = new System.Drawing.Size(600, 20)
            };
            this.Controls.Add(txtConnectionString);

            // Order
            var lblOrder = new Label
            {
                Text = "Order:",
                Location = new System.Drawing.Point(20, 400),
                Size = new System.Drawing.Size(100, 20)
            };
            this.Controls.Add(lblOrder);

            numConnectionOrder = new NumericUpDown
            {
                Location = new System.Drawing.Point(130, 398),
                Size = new System.Drawing.Size(80, 20),
                Minimum = 0,
                Maximum = 1000,
                Value = 0
            };
            this.Controls.Add(numConnectionOrder);

            // Buttons
            btnAddConnection = new Button
            {
                Text = "Add",
                Location = new System.Drawing.Point(20, 440),
                Size = new System.Drawing.Size(80, 30)
            };
            btnAddConnection.Click += BtnAddConnection_Click;
            this.Controls.Add(btnAddConnection);

            btnUpdateConnection = new Button
            {
                Text = "Update",
                Location = new System.Drawing.Point(110, 440),
                Size = new System.Drawing.Size(80, 30)
            };
            btnUpdateConnection.Click += BtnUpdateConnection_Click;
            this.Controls.Add(btnUpdateConnection);

            btnDeleteConnection = new Button
            {
                Text = "Delete",
                Location = new System.Drawing.Point(200, 440),
                Size = new System.Drawing.Size(80, 30)
            };
            btnDeleteConnection.Click += BtnDeleteConnection_Click;
            this.Controls.Add(btnDeleteConnection);

            btnTestSelectedConnection = new Button
            {
                Text = "Test Connection",
                Location = new System.Drawing.Point(290, 440),
                Size = new System.Drawing.Size(120, 30)
            };
            btnTestSelectedConnection.Click += BtnTestSelectedConnection_Click;
            this.Controls.Add(btnTestSelectedConnection);

            // Initialize DataGridView columns
            dgvConnections.Columns.Add("Id", "ID");
            dgvConnections.Columns.Add("Name", "Name");
            dgvConnections.Columns.Add("ConnectionString", "Connection String");
            dgvConnections.Columns.Add("Order", "Order");
            
            dgvConnections.Columns["Id"].Visible = false;
            dgvConnections.Columns["ConnectionString"].Width = 300;
            dgvConnections.Columns["Order"].Width = 60;

            this.ResumeLayout(false);
        }

        // Public methods để MainForm có thể gọi
        public void SetConfig(AppConfig sharedConfig)
        {
            config = sharedConfig;
            RefreshConnectionsList();
        }

        public void LoadConfiguration()
        {
            try
            {
                config = AppConfig.Load();
                RefreshConnectionsList();
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
                config.Save();
                FileLogger.LogInfo("Configuration saved");
            }
            catch (Exception ex)
            {
                LogMessage($"Lỗi lưu cấu hình: {ex.Message}");
                FileLogger.LogError("Error saving configuration", ex);
            }
        }

        // Event Handlers
        private void DgvConnections_SelectionChanged(object sender, EventArgs e)
        {
            if (dgvConnections.SelectedRows.Count > 0)
            {
                var row = dgvConnections.SelectedRows[0];
                txtConnectionName.Text = row.Cells["Name"].Value?.ToString() ?? "";
                txtConnectionString.Text = row.Cells["ConnectionString"].Value?.ToString() ?? "";
                numConnectionOrder.Value = Convert.ToDecimal(row.Cells["Order"].Value ?? 0);
            }
        }

        private void BtnAddConnection_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtConnectionName.Text) || string.IsNullOrWhiteSpace(txtConnectionString.Text))
            {
                MessageBox.Show("Vui lòng nhập đầy đủ tên và chuỗi kết nối!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var newConnection = new ConnectionStringItem
            {
                Id = Guid.NewGuid().ToString(),
                Name = txtConnectionName.Text.Trim(),
                ConnectionString = txtConnectionString.Text.Trim(),
                Order = (int)numConnectionOrder.Value
            };

            config.ConnectionStrings.Add(newConnection);
            config.ConnectionStrings = config.ConnectionStrings.OrderBy(x => x.Order).ToList();
            
            RefreshConnectionsList();
            ClearConnectionForm();
            LogMessage($"Đã thêm connection: {newConnection.Name}");
            
            // Trigger configuration save event to notify MainForm
            ConfigurationSaveRequested?.Invoke(this, config);
        }

        private void BtnUpdateConnection_Click(object sender, EventArgs e)
        {
            if (dgvConnections.SelectedRows.Count == 0)
            {
                MessageBox.Show("Vui lòng chọn connection cần cập nhật!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (string.IsNullOrWhiteSpace(txtConnectionName.Text) || string.IsNullOrWhiteSpace(txtConnectionString.Text))
            {
                MessageBox.Show("Vui lòng nhập đầy đủ tên và chuỗi kết nối!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var row = dgvConnections.SelectedRows[0];
            var connectionId = row.Cells["Id"].Value?.ToString();
            var connection = config.ConnectionStrings.FirstOrDefault(x => x.Id == connectionId);
            
            if (connection != null)
            {
                connection.Name = txtConnectionName.Text.Trim();
                connection.ConnectionString = txtConnectionString.Text.Trim();
                connection.Order = (int)numConnectionOrder.Value;
                
                config.ConnectionStrings = config.ConnectionStrings.OrderBy(x => x.Order).ToList();
                
                RefreshConnectionsList();
                LogMessage($"Đã cập nhật connection: {connection.Name}");
                
                // Trigger configuration save event to notify MainForm
                ConfigurationSaveRequested?.Invoke(this, config);
            }
        }

        private void BtnDeleteConnection_Click(object sender, EventArgs e)
        {
            if (dgvConnections.SelectedRows.Count == 0)
            {
                MessageBox.Show("Vui lòng chọn connection cần xóa!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var result = MessageBox.Show("Bạn có chắc chắn muốn xóa connection này?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
            {
                var row = dgvConnections.SelectedRows[0];
                var connectionId = row.Cells["Id"].Value?.ToString();
                var connection = config.ConnectionStrings.FirstOrDefault(x => x.Id == connectionId);
                
                if (connection != null)
                {
                    config.ConnectionStrings.Remove(connection);
                    
                    RefreshConnectionsList();
                    ClearConnectionForm();
                    
                    LogMessage($"Đã xóa connection: {connection.Name}");
                    
                    // Trigger configuration save event to notify MainForm
                    ConfigurationSaveRequested?.Invoke(this, config);
                }
            }
        }

        private void BtnTestSelectedConnection_Click(object sender, EventArgs e)
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
                    LogMessage($"✓ Test connection thành công: {txtConnectionName.Text}");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi kết nối: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                LogMessage($"✗ Test connection thất bại: {ex.Message}");
            }
        }

        // Helper methods
        private void RefreshConnectionsList()
        {
            if (InvokeRequired)
            {
                Invoke(new Action(RefreshConnectionsList));
                return;
            }

            dgvConnections.Rows.Clear();
            foreach (var connection in config.ConnectionStrings.OrderBy(x => x.Order))
            {
                dgvConnections.Rows.Add(connection.Id, connection.Name, connection.ConnectionString, connection.Order);
            }
        }

        private void ClearConnectionForm()
        {
            txtConnectionName.Text = "";
            txtConnectionString.Text = "";
            numConnectionOrder.Value = 0;
        }

        private void LogMessage(string message)
        {
            LogMessageRequested?.Invoke(this, message);
        }
    }
}
