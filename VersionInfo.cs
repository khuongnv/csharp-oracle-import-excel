using System;

namespace ExcelToOracleImporter
{
    public static class VersionInfo
    {
        public const string Version = "2.1.1";
        public const string BuildDate = "2025-09-29";
        public const string LastUpdate = "2025-09-29";
        
        public static string GetFullVersion()
        {
            return $"v{Version} (Build {BuildDate})";
        }
        
        public static string GetChangelog()
        {
            return @"CHANGELOG - Excel to Oracle Database Importer
================================================

Version 2.1.1 (2025-09-29)
---------------------------
✨ NEW FEATURES:
• Added Log menu in Help menu to display changelog
• Added VersionInfo class for centralized version management
• Added LogForm for displaying application changelog with version history

🔧 IMPROVEMENTS:
• Enhanced Help menu with Log submenu
• Improved version display in application title
• Better changelog formatting and readability
• Centralized version information management

Version 2.1.0 (2025-09-29)
---------------------------
✨ NEW FEATURES:
• Refactored to UserControl architecture for better maintainability
• Converted from TabControl to MenuStrip navigation
• Added Connection Management with CRUD operations
• Real-time connection list refresh between menus

🔧 IMPROVEMENTS:
• Split MainForm into ImportTab and ConnectionManagementTab UserControls
• Added automatic refresh of connection list when changes are made
• Improved code organization and separation of concerns
• Better error handling and logging
• Enhanced menu structure with Help submenu

🐛 BUG FIXES:
• Fixed ORA-00911 invalid character error with proper Oracle naming conventions
• Fixed Vietnamese character conversion for table and column names
• Fixed column comment functionality
• Fixed connection string management synchronization
• Fixed connection list not refreshing after CRUD operations

Version 2.0.0 (2025-09-28)
---------------------------
✨ NEW FEATURES:
• Multi-sheet Excel support - each sheet creates separate Oracle table
• Sheet selection option (single sheet or all sheets)
• Column comments in Oracle tables (original Excel header names)
• Connection string management with CRUD operations
• Vietnamese character conversion for proper Oracle naming

🔧 IMPROVEMENTS:
• Enhanced table naming: single sheet uses input name, multiple sheets use input_name_SheetX
• Better column name cleaning with Vietnamese character support
• Improved error handling and user feedback
• Added comprehensive logging system

🐛 BUG FIXES:
• Fixed Oracle identifier naming issues
• Fixed special character handling in table/column names
• Fixed batch processing for large datasets

Version 1.0.0 (2025-09-27)
---------------------------
🎉 INITIAL RELEASE:
• Basic Excel to Oracle import functionality
• Support for .xlsx and .xls files
• Batch processing for large datasets
• Connection testing
• Data preview functionality
• Basic logging system

TECHNICAL DETAILS:
==================
• Framework: .NET 6.0 Windows Forms
• Database: Oracle (using Oracle.ManagedDataAccess.Client)
• Excel Processing: EPPlus
• Configuration: JSON-based settings
• Logging: File-based logging system

ARCHITECTURE:
=============
• MainForm: Main application form with MenuStrip navigation
• ImportTab: UserControl for Excel import functionality
• ConnectionManagementTab: UserControl for connection string management
• AppConfig: Configuration management with JSON serialization
• FileLogger: Logging system for application events
• AboutForm: Application information dialog
• LogForm: Changelog display dialog

DEVELOPMENT NOTES:
==================
• Code is organized using UserControl pattern for better maintainability
• Event-driven architecture for communication between components
• Comprehensive error handling and logging throughout the application
• Vietnamese language support for user interface and data processing
• Oracle naming conventions compliance for table and column names

For technical support or feature requests, please contact the development team.
";
        }
    }
}
