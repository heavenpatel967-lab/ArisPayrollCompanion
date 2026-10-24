using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using Microsoft.Data.Sqlite;
using ArisPayrollCompanion.Helpers;

namespace ArisPayrollCompanion.Views
{
    public partial class ComplianceDashboardWindow : Window
    {
        public ComplianceDashboardWindow()
        {
            InitializeComponent();
            LoadComplianceData();
        }

        private void LoadComplianceData()
        {
            List<ComplianceRecord> auditList = new List<ComplianceRecord>();
            int missingPhotos = 0;
            int missingAadhar = 0;
            int missingVillage = 0;

            try
            {
                using (var connection = new SqliteConnection(DatabaseHelper.GetConnectionString()))
                {
                    connection.Open();
                    // Pull all profiles currently in the vault
                    string query = "SELECT EmployeeCode, OriginalName, AadharNumber, NativePlace, Village, PhotoFilePath FROM EmployeeProfiles";

                    using (var command = new SqliteCommand(query, connection))
                    {
                        using (var reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                string empCode = reader["EmployeeCode"].ToString();
                                string name = reader["OriginalName"].ToString();
                                if (string.IsNullOrWhiteSpace(name)) name = "Name Not Saved";

                                string aadhar = reader["AadharNumber"].ToString();
                                string nativePlace = reader["NativePlace"].ToString();
                                string village = reader["Village"].ToString();
                                string photoPath = reader["PhotoFilePath"].ToString();

                                // Evaluate Data
                                bool hasPhoto = !string.IsNullOrWhiteSpace(photoPath) && File.Exists(photoPath);
                                bool hasAadhar = !string.IsNullOrWhiteSpace(aadhar);
                                bool hasVillage = !string.IsNullOrWhiteSpace(nativePlace) || !string.IsNullOrWhiteSpace(village);

                                // Increment Top Counters
                                if (!hasPhoto) missingPhotos++;
                                if (!hasAadhar) missingAadhar++;
                                if (!hasVillage) missingVillage++;

                                int flagCount = (!hasPhoto ? 1 : 0) + (!hasAadhar ? 1 : 0) + (!hasVillage ? 1 : 0);

                                // Only add to the grid if they are missing at least one thing
                                if (flagCount > 0)
                                {
                                    auditList.Add(new ComplianceRecord
                                    {
                                        EmployeeCode = empCode,
                                        EmployeeName = name,
                                        PhotoStatus = hasPhoto ? "OK" : "MISSING",
                                        PhotoColor = hasPhoto ? "#27AE60" : "#E74C3C",
                                        AadharStatus = hasAadhar ? "OK" : "MISSING",
                                        AadharColor = hasAadhar ? "#27AE60" : "#E74C3C",
                                        VillageStatus = hasVillage ? "OK" : "MISSING",
                                        VillageColor = hasVillage ? "#27AE60" : "#E74C3C",
                                        FlagCount = flagCount
                                    });
                                }
                            }
                        }
                    }
                }

                // Update UI Counters
                txtMissingPhotos.Text = missingPhotos.ToString();
                txtMissingAadhar.Text = missingAadhar.ToString();
                txtMissingVillage.Text = missingVillage.ToString();

                // Sort the list so people missing the MOST data show up at the top
                dgCompliance.ItemsSource = auditList.OrderByDescending(x => x.FlagCount).ToList();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error running compliance audit: {ex.Message}", "System Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }

    // Helper class to bind data to the grid
    public class ComplianceRecord
    {
        public string EmployeeCode { get; set; }
        public string EmployeeName { get; set; }
        public string PhotoStatus { get; set; }
        public string PhotoColor { get; set; }
        public string AadharStatus { get; set; }
        public string AadharColor { get; set; }
        public string VillageStatus { get; set; }
        public string VillageColor { get; set; }
        public int FlagCount { get; set; }
    }
}