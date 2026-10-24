using Microsoft.Data.Sqlite;
using System;
using System.IO;

namespace ArisPayrollCompanion.Database
{
    public class DatabaseService
    {
        // Points directly to the root folder where your installer drops the .exe and Payroll.db
        private static readonly string DatabaseFile =
            Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Payroll.db");

        private static readonly string ConnectionString =
            $"Data Source={DatabaseFile}";

        public void InitializeDatabase()
        {
            // SQLite automatically creates the DB file if it doesn't exist when opened,
            // but this will now correctly target the pre-packaged file from your installer.
            using var connection = new SqliteConnection(ConnectionString);
            connection.Open();
            connection.Close();
        }

        public string GetDatabasePath()
        {
            return DatabaseFile;
        }

        public string GetConnectionString()
        {
            return ConnectionString;
        }
    }
}