using System;
using System.IO;
using Microsoft.Data.Sqlite;
using System.Windows;

namespace ArisPayrollCompanion.Helpers
{
    public static class DatabaseHelper
    {
        // DYNAMIC PATH: This targets the exact folder where the app (.exe) is running on ANY computer
        private static readonly string AppFolder = AppDomain.CurrentDomain.BaseDirectory;

        // Updated to match the actual file we packaged into your installer!
        private static readonly string DbPath = Path.Combine(AppFolder, "Payroll.db");

        public static string GetConnectionString()
        {
            // Mode=ReadWriteCreate ensures it builds a fresh database if one is missing,
            // but it will seamlessly connect to your pre-packaged Payroll.db first.
            return $"Data Source={DbPath};Mode=ReadWriteCreate;";
        }

        public static void InitializeDatabase()
        {
            try
            {
                using (var connection = new SqliteConnection(GetConnectionString()))
                {
                    connection.Open();
                    using (var command = connection.CreateCommand())
                    {
                        command.CommandText = @"
                            CREATE TABLE IF NOT EXISTS EmployeeProfiles (
                                EmployeeCode TEXT PRIMARY KEY, OriginalName TEXT, AadharName TEXT, 
                                AadharNumber TEXT, PhoneNumber TEXT, NativePlace TEXT, 
                                Village TEXT, PhotoFilePath TEXT, Remarks TEXT, DutiesAssigned TEXT
                            );
                            CREATE TABLE IF NOT EXISTS LongLeaveTracking (
                                LeaveID INTEGER PRIMARY KEY AUTOINCREMENT, EmployeeCode TEXT, 
                                DepartureDate TEXT, ExpectedReturnDate TEXT, ActualReturnDate TEXT, 
                                LateDays INTEGER, Reason TEXT
                            );
                            CREATE TABLE IF NOT EXISTS EmployeeDocuments (
                                DocID INTEGER PRIMARY KEY AUTOINCREMENT, EmployeeCode TEXT, 
                                DocumentCategory TEXT, DocumentName TEXT, FilePath TEXT, UploadDate TEXT
                            );";
                        command.ExecuteNonQuery();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Database Init Error: {ex.Message}", "Fatal Error");
            }
        }
    }
}