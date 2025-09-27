using System;
using System.IO;

namespace ExcelToOracleImporter
{
    public static class FileLogger
    {
        private static readonly string LogPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "logs");
        private static readonly string LogFile = Path.Combine(LogPath, $"import_log_{DateTime.Now:yyyyMMdd}.txt");

        static FileLogger()
        {
            // Tạo thư mục logs nếu chưa tồn tại
            if (!Directory.Exists(LogPath))
            {
                Directory.CreateDirectory(LogPath);
            }
        }

        public static void Log(string message)
        {
            try
            {
                var logEntry = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {message}{Environment.NewLine}";
                File.AppendAllText(LogFile, logEntry);
            }
            catch (Exception ex)
            {
                // Nếu không ghi được log file, chỉ log ra console
                System.Diagnostics.Debug.WriteLine($"Error writing to log file: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"Original message: {message}");
            }
        }

        public static void LogError(string message, Exception? exception = null)
        {
            var errorMessage = $"ERROR: {message}";
            if (exception != null)
            {
                errorMessage += $" - Exception: {exception.Message}";
            }
            Log(errorMessage);
        }

        public static void LogInfo(string message)
        {
            Log($"INFO: {message}");
        }

        public static void LogSuccess(string message)
        {
            Log($"SUCCESS: {message}");
        }

        public static string GetLogFilePath()
        {
            return LogFile;
        }
    }
}
