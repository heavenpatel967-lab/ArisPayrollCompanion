using System;
using System.IO;
using System.Windows;
using ArisPayrollCompanion.Helpers;

namespace ArisPayrollCompanion
{
    public partial class App : System.Windows.Application
    {
        private static readonly string ConfigFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "ArisPayrollCompanion");
        private static readonly string ConfigFile = Path.Combine(ConfigFolder, "DbfPathConfig.txt");

        private static string _dbfFolderPath;

        // Triggers the moment the app opens to ensure all missing SQLite tables are created instantly
        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);
            DatabaseHelper.InitializeDatabase();
        }

        public static string DbfFolderPath
        {
            get
            {
                if (_dbfFolderPath == null)
                {
                    if (File.Exists(ConfigFile))
                    {
                        _dbfFolderPath = File.ReadAllText(ConfigFile).Trim();
                    }
                    else
                    {
                        _dbfFolderPath = @"C:\";
                    }
                }
                return _dbfFolderPath;
            }
            set
            {
                _dbfFolderPath = value;
                Directory.CreateDirectory(ConfigFolder);
                File.WriteAllText(ConfigFile, value);
            }
        }
    }
}