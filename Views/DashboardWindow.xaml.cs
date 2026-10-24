using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using Microsoft.Data.Sqlite;
using ArisPayrollCompanion.Helpers;

namespace ArisPayrollCompanion.Views
{
    public partial class DashboardWindow : Window
    {
        private readonly string _dbfFolderPath;

        public DashboardWindow(string dbfFolderPath)
        {
            InitializeComponent();
            _dbfFolderPath = dbfFolderPath;

            LoadFinancialAndHeadcountData();
            LoadLateReturnAlerts();
        }

        private void LoadFinancialAndHeadcountData()
        {
            try
            {
                DataTable dt = ReadDbfManually("monsum.DBF");
                if (dt == null || dt.Rows.Count == 0)
                {
                    txtTotalEmployees.Text = "0";
                    txtGlobalLoan.Text = "₹ 0.00";
                    txtGlobalAdvance.Text = "₹ 0.00";
                    return;
                }

                // 1. Calculate Active Headcount (Unique Employee Codes in DBF)
                var allEmpCodes = dt.AsEnumerable()
                    .Where(r => dt.Columns.Contains("EMPCD"))
                    .Select(r => r["EMPCD"].ToString().Trim())
                    .Where(c => !string.IsNullOrEmpty(c))
                    .Distinct()
                    .ToList();

                txtTotalEmployees.Text = allEmpCodes.Count.ToString("N0");

                // 2. Calculate Factory-Wide Loan
                // Group by EMPCD so we only calculate the absolute latest loan balance per employee
                var latestLoanRecords = dt.AsEnumerable()
                    .Where(r => dt.Columns.Contains("EMPCD") && dt.Columns.Contains("YRMONTH"))
                    .GroupBy(r => r["EMPCD"].ToString().Trim())
                    .Select(g => g.OrderBy(r => r["YRMONTH"].ToString()).Last()) // Get latest month for each employee
                    .ToList();

                double totalFactoryLoan = 0;
                foreach (var row in latestLoanRecords)
                {
                    double paidLoan = dt.Columns.Contains("PAID_LOAN") ? ParseAmount(row["PAID_LOAN"]) : 0;
                    double loanRec = dt.Columns.Contains("LOAN_REC") ? ParseAmount(row["LOAN_REC"]) : 0;

                    double outstanding = paidLoan - loanRec;
                    if (outstanding > 0)
                    {
                        totalFactoryLoan += outstanding;
                    }
                }

                // 3. Calculate Factory-Wide Advances (Sum of all advances given in the ledger)
                double totalFactoryAdvance = dt.AsEnumerable()
                    .Where(r => dt.Columns.Contains("ADVANCE"))
                    .Sum(r => ParseAmount(r["ADVANCE"]));

                // Update UI
                txtGlobalLoan.Text = $"₹ {totalFactoryLoan:N2}";
                txtGlobalAdvance.Text = $"₹ {totalFactoryAdvance:N2}";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error calculating dashboard metrics: {ex.Message}");
            }
        }

        private void LoadLateReturnAlerts()
        {
            List<LateReturnRecord> alertList = new List<LateReturnRecord>();
            DateTime today = DateTime.Today;

            try
            {
                using (var connection = new SqliteConnection(DatabaseHelper.GetConnectionString()))
                {
                    connection.Open();
                    // Join Leaves with Profiles to get the employee's name
                    string query = @"
                        SELECT l.EmployeeCode, p.OriginalName, l.DepartureDate, l.ExpectedReturnDate, l.Reason 
                        FROM LongLeaveTracking l 
                        LEFT JOIN EmployeeProfiles p ON l.EmployeeCode = p.EmployeeCode 
                        WHERE l.ActualReturnDate = '' OR l.ActualReturnDate IS NULL";

                    using (var command = new SqliteCommand(query, connection))
                    {
                        using (var reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                string expectedDateStr = reader["ExpectedReturnDate"].ToString();

                                if (DateTime.TryParse(expectedDateStr, out DateTime expectedDate))
                                {
                                    // Check if they are officially late
                                    if (expectedDate < today)
                                    {
                                        int daysLate = (today - expectedDate).Days;

                                        string name = reader["OriginalName"].ToString();
                                        if (string.IsNullOrWhiteSpace(name)) name = "Name Not Saved in Vault";

                                        alertList.Add(new LateReturnRecord
                                        {
                                            EmployeeCode = reader["EmployeeCode"].ToString(),
                                            EmployeeName = name,
                                            DepartureDate = reader["DepartureDate"].ToString(),
                                            ExpectedReturnDate = expectedDateStr,
                                            Reason = reader["Reason"].ToString(),
                                            DaysLate = $"+{daysLate} Days"
                                        });
                                    }
                                }
                            }
                        }
                    }
                }

                // Bind to Grid, sorted with the most late people at the top
                dgLateReturns.ItemsSource = alertList.OrderByDescending(a => int.Parse(a.DaysLate.Replace("+", "").Replace(" Days", ""))).ToList();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading late returns: {ex.Message}");
            }
        }

        // --- Helper Methods ---
        private DataTable ReadDbfManually(string fileName)
        {
            DataTable dt = new DataTable();
            string filePath = Path.Combine(_dbfFolderPath, fileName);
            if (!File.Exists(filePath)) return dt;

            using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var reader = new BinaryReader(stream, Encoding.Latin1))
            {
                reader.ReadBytes(4);
                int numRecords = reader.ReadInt32();
                int headerSize = reader.ReadInt16();
                int recordSize = reader.ReadInt16();
                reader.ReadBytes(20);

                List<Tuple<string, int>> fields = new List<Tuple<string, int>>();
                while (reader.BaseStream.Position < headerSize - 1)
                {
                    byte[] nameBytes = reader.ReadBytes(11);
                    string name = Encoding.ASCII.GetString(nameBytes).Replace("\0", "").Trim();
                    reader.ReadByte(); reader.ReadBytes(4);
                    int len = reader.ReadByte();
                    reader.ReadBytes(15);
                    if (!string.IsNullOrEmpty(name))
                    {
                        fields.Add(new Tuple<string, int>(name, len));
                        dt.Columns.Add(name);
                    }
                }
                reader.BaseStream.Position = headerSize;
                for (int i = 0; i < numRecords; i++)
                {
                    byte[] rowData = reader.ReadBytes(recordSize);
                    if (rowData.Length < recordSize) break;
                    DataRow row = dt.NewRow();
                    int offset = 1;
                    foreach (var field in fields)
                    {
                        if (offset + field.Item2 <= rowData.Length)
                            row[field.Item1] = Encoding.Latin1.GetString(rowData, offset, field.Item2).Trim();
                        offset += field.Item2;
                    }
                    dt.Rows.Add(row);
                }
            }
            return dt;
        }

        private double ParseAmount(object value)
        {
            if (value == null) return 0;
            string s = value.ToString().Trim();
            if (string.IsNullOrEmpty(s)) return 0;
            return double.TryParse(s, out double result) ? result : 0;
        }
    }

    public class LateReturnRecord
    {
        public string EmployeeCode { get; set; }
        public string EmployeeName { get; set; }
        public string DepartureDate { get; set; }
        public string ExpectedReturnDate { get; set; }
        public string DaysLate { get; set; }
        public string Reason { get; set; }
    }
}