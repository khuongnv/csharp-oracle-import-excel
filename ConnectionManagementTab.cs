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
            this.Resize += ConnectionManagementTab_Resize;
            this.Load += ConnectionManagementTab_Load;
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();

            // Connection Name
            var lblConnectionName = new Label
            {
                Text = "Name:",
                Location = new System.Drawing.Point(20, 20),
                Size = new System.Drawing.Size(100, 20),
                Anchor = AnchorStyles.Top | AnchorStyles.Left
            };
            this.Controls.Add(lblConnectionName);

            txtConnectionName = new TextBox
            {
                Location = new System.Drawing.Point(130, 18),
                Size = new System.Drawing.Size(200, 20),
                Anchor = AnchorStyles.Top | AnchorStyles.Left
            };
            this.Controls.Add(txtConnectionName);

            // Connection String
            var lblConnectionString = new Label
            {
                Text = "Connection:",
                Location = new System.Drawing.Point(20, 50),
                Size = new System.Drawing.Size(100, 20),
                Anchor = AnchorStyles.Top | AnchorStyles.Left
            };
            this.Controls.Add(lblConnectionString);

            txtConnectionString = new TextBox
            {
                Location = new System.Drawing.Point(130, 48),
                Size = new System.Drawing.Size(600, 20),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            this.Controls.Add(txtConnectionString);

            // Order
            var lblOrder = new Label
            {
                Text = "Order:",
                Location = new System.Drawing.Point(20, 80),
                Size = new System.Drawing.Size(100, 20),
                Anchor = AnchorStyles.Top | AnchorStyles.Left
            };
            this.Controls.Add(lblOrder);

            numConnectionOrder = new NumericUpDown
            {
                Location = new System.Drawing.Point(130, 78),
                Size = new System.Drawing.Size(80, 20),
                Minimum = 0,
                Maximum = 1000,
                Value = 0,
                Anchor = AnchorStyles.Top | AnchorStyles.Left
            };
            this.Controls.Add(numConnectionOrder);

            // Buttons
            btnAddConnection = new Button
            {
                Text = "Add",
                Location = new System.Drawing.Point(20, 120),
                Size = new System.Drawing.Size(80, 30),
                Anchor = AnchorStyles.Top | AnchorStyles.Left
            };
            btnAddConnection.Click += BtnAddConnection_Click;
            this.Controls.Add(btnAddConnection);

            btnUpdateConnection = new Button
            {
                Text = "Update",
                Location = new System.Drawing.Point(110, 120),
                Size = new System.Drawing.Size(80, 30),
                Anchor = AnchorStyles.Top | AnchorStyles.Left
            };
            btnUpdateConnection.Click += BtnUpdateConnection_Click;
            this.Controls.Add(btnUpdateConnection);

            btnDeleteConnection = new Button
            {
                Text = "Delete",
                Location = new System.Drawing.Point(200, 120),
                Size = new System.Drawing.Size(80, 30),
                Anchor = AnchorStyles.Top | AnchorStyles.Left
            };
            btnDeleteConnection.Click += BtnDeleteConnection_Click;
            this.Controls.Add(btnDeleteConnection);

            btnTestSelectedConnection = new Button
            {
                Text = "Test Connection",
                Location = new System.Drawing.Point(290, 120),
                Size = new System.Drawing.Size(120, 30),
                Anchor = AnchorStyles.Top | AnchorStyles.Left
            };
            btnTestSelectedConnection.Click += BtnTestSelectedConnection_Click;
            this.Controls.Add(btnTestSelectedConnection);

            // DataGridView for connections
            dgvConnections = new DataGridView
            {
                Location = new System.Drawing.Point(20, 170),
                Size = new System.Drawing.Size(960, 500),
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                ReadOnly = true,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false,
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            };
            dgvConnections.SelectionChanged += DgvConnections_SelectionChanged;
            this.Controls.Add(dgvConnections);

            // Initialize DataGridView columns
            dgvConnections.Columns.Add("Id", "ID");
            dgvConnections.Columns.Add("Name", "Name");
            dgvConnections.Columns.Add("ConnectionString", "Connection String");
            dgvConnections.Columns.Add("Order", "Order");
            
            dgvConnections.Columns["Id"].Visible = false;
            dgvConnections.Columns["Name"].Width = 200;
            dgvConnections.Columns["ConnectionString"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dgvConnections.Columns["Order"].Width = 80;

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

        private void ConnectionManagementTab_Resize(object sender, EventArgs e)
        {
            if (dgvConnections != null)
            {
                dgvConnections.Size = new System.Drawing.Size(this.Width - 40, this.Height - 200);
            }
            
            if (txtConnectionString != null)
            {
                txtConnectionString.Size = new System.Drawing.Size(this.Width - 150, 20);
            }
        }

        private void ConnectionManagementTab_Load(object sender, EventArgs e)
        {
            // Ensure proper layout when the control is loaded
            if (dgvConnections != null)
            {
                dgvConnections.Size = new System.Drawing.Size(this.Width - 40, this.Height - 200);
            }
            
            if (txtConnectionString != null)
            {
                txtConnectionString.Size = new System.Drawing.Size(this.Width - 150, 20);
            }
        }
    }
}
