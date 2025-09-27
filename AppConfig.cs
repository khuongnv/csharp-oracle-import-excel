using System;
using System.IO;
using Newtonsoft.Json;

namespace ExcelToOracleImporter
{
    public class AppConfig
    {
        public string ConnectionString { get; set; } = "";
        public string TableName { get; set; } = "EXCEL_IMPORT";
        public bool HasHeader { get; set; } = true;
        public int BatchSize { get; set; } = 100;
        public string LastExcelFilePath { get; set; } = "";

        private static readonly string ConfigPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "config.json");

        public static AppConfig Load()
        {
            try
            {
                if (File.Exists(ConfigPath))
                {
                    var json = File.ReadAllText(ConfigPath);
                    return JsonConvert.DeserializeObject<AppConfig>(json) ?? new AppConfig();
                }
            }
            catch (Exception ex)
            {
                // Log error but don't crash the app
                System.Diagnostics.Debug.WriteLine($"Error loading config: {ex.Message}");
            }
            return new AppConfig();
        }

        public void Save()
        {
            try
            {
                var json = JsonConvert.SerializeObject(this, Formatting.Indented);
                File.WriteAllText(ConfigPath, json);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error saving config: {ex.Message}");
            }
        }
    }
}
