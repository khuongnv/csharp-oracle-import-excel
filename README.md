# Excel to Oracle Database Importer

Ứng dụng Windows Forms để import dữ liệu từ file Excel (.xlsx, .xls) vào Oracle Database.

## Tính năng

- ✅ Import dữ liệu từ Excel vào Oracle Database
- ✅ Hỗ trợ file có header hoặc không có header
- ✅ Tự động tạo bảng trong Oracle với tên cột phù hợp
- ✅ Preview dữ liệu trước khi import
- ✅ Batch processing với progress bar
- ✅ Test kết nối Oracle database
- ✅ Log chi tiết quá trình import
- ✅ Xử lý lỗi và validation đầu vào
- ✅ **Lưu cấu hình tự động** - Connection string và settings được lưu vào JSON
- ✅ **File logging** - Ghi log ra file text với timestamp
- ✅ **Mở thư mục logs** - Nút để mở thư mục chứa log files

## Yêu cầu hệ thống

- Windows 10/11
- .NET 6.0 Runtime
- Oracle Database (local hoặc remote)
- File Excel (.xlsx hoặc .xls)

## Cài đặt và chạy

### Cách 1: Sử dụng Visual Studio 2022 (Khuyến nghị)
1. **Clone repository**
   ```bash
   git clone https://github.com/khuongnv/csharp-oracle-import-excel.git
   cd csharp-oracle-import-excel
   ```

2. **Mở solution**
   - Mở Visual Studio 2022
   - File → Open → Project/Solution
   - Chọn file `ExcelToOracleImporter.sln`

3. **Restore packages**
   - Visual Studio sẽ tự động restore packages
   - Hoặc nhấn chuột phải vào Solution → Restore NuGet Packages

4. **Build và chạy**
   - Nhấn F5 hoặc Ctrl+F5 để chạy
   - Hoặc Build → Build Solution

### Cách 2: Sử dụng .NET CLI
1. **Clone repository**
   ```bash
   git clone https://github.com/khuongnv/csharp-oracle-import-excel.git
   cd csharp-oracle-import-excel
   ```

2. **Restore packages**
   ```bash
   dotnet restore
   ```

3. **Build project**
   ```bash
   dotnet build
   ```

4. **Chạy ứng dụng**
   ```bash
   dotnet run
   ```

## Hướng dẫn sử dụng

### 1. Cấu hình kết nối Oracle
- Nhập Connection String của Oracle Database
- Format: `Data Source=hostname:port/service_name;User Id=username;Password=password;`
- Ví dụ: `Data Source=localhost:1521/XE;User Id=hr;Password=password;`
- Nhấn "Test Connection" để kiểm tra kết nối

### 2. Chọn file Excel
- Nhấn "Select Excel File" để chọn file Excel
- Hỗ trợ format .xlsx và .xls
- Nhấn "Preview Data" để xem trước dữ liệu

### 3. Cấu hình import
- **Table Name**: Tên bảng sẽ được tạo trong Oracle (mặc định: EXCEL_IMPORT)
- **Has Header**: Check nếu dòng đầu tiên là tên cột
- **Batch Size**: Số dòng xử lý mỗi batch (mặc định: 100)

### 4. Import dữ liệu
- Nhấn "Import Data" để bắt đầu import
- Theo dõi progress bar và log
- Ứng dụng sẽ tự động:
  - Tạo bảng mới (xóa bảng cũ nếu tồn tại)
  - Import tất cả dữ liệu từ Excel
  - Hiển thị kết quả

## Cấu trúc dữ liệu

### Tên cột
- Nếu có header: Sử dụng tên cột từ Excel (được làm sạch để phù hợp Oracle)
- Nếu không có header: Tạo tên cột tự động (A, B, C, ..., AA, AB, ...)

### Kiểu dữ liệu
- Tất cả cột được tạo dưới dạng `VARCHAR2(4000)`
- Dữ liệu được chuyển đổi thành string

## Xử lý lỗi thường gặp

### 1. Lỗi kết nối Oracle
```
ORA-12154: TNS:could not resolve the connect identifier
```
**Giải pháp**: Kiểm tra Connection String và tên service

### 2. Lỗi quyền truy cập
```
ORA-00942: table or view does not exist
```
**Giải pháp**: Đảm bảo user có quyền CREATE TABLE

### 3. Lỗi đọc file Excel
```
File Excel không có dữ liệu!
```
**Giải pháp**: Kiểm tra file Excel có dữ liệu và đúng format

### 4. Lỗi tên cột không hợp lệ
**Giải pháp**: Ứng dụng tự động làm sạch tên cột, loại bỏ ký tự đặc biệt

## Packages sử dụng

- `Oracle.ManagedDataAccess.Core` (23.6.0) - Kết nối Oracle
- `EPPlus` (7.0.5) - Đọc file Excel
- `System.Data.OleDb` (7.0.0) - Hỗ trợ Excel cũ
- `Newtonsoft.Json` (13.0.3) - Xử lý JSON cho cấu hình

## Phát triển

### Cấu trúc project
```
ExcelToOracleImporter/
├── ExcelToOracleImporter.sln    # Visual Studio Solution file
├── ExcelToOracleImporter.csproj # Project file
├── MainForm.cs                  # Giao diện chính
├── Program.cs                   # Entry point
├── AppConfig.cs                 # Quản lý cấu hình JSON
├── FileLogger.cs                # Ghi log ra file
├── .gitignore                   # Git ignore file
├── config.json                  # File cấu hình (tự tạo)
├── logs/                        # Thư mục chứa log files
└── README.md                   # Hướng dẫn này
```

### Tính năng mới v2.0

#### 📁 **Lưu cấu hình tự động**
- Connection string, table name, batch size, header settings được lưu vào `config.json`
- Tự động load cấu hình khi khởi động ứng dụng
- Không cần nhập lại thông tin mỗi lần sử dụng

#### 📝 **File Logging**
- Ghi log chi tiết ra file text trong thư mục `logs/`
- Format log: `import_log_YYYYMMDD.txt`
- Log bao gồm: INFO, SUCCESS, ERROR với timestamp
- Nút "Open Logs Folder" để mở thư mục logs

#### 🔧 **Cấu hình JSON**
```json
{
  "ConnectionString": "Data Source=localhost:1521/XE;User Id=hr;Password=password;",
  "TableName": "EXCEL_IMPORT",
  "HasHeader": true,
  "BatchSize": 100,
  "LastExcelFilePath": "C:\\path\\to\\file.xlsx"
}
```

### Build release
```bash
dotnet build -c Release
```

### Publish standalone
```bash
dotnet publish -c Release -r win-x64 --self-contained
```

## License

MIT License

## Hỗ trợ

Nếu gặp vấn đề, vui lòng tạo issue hoặc liên hệ developer.
