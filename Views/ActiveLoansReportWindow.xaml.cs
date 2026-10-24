using ArisPayrollCompanion.Helpers;
using ArisPayrollCompanion.Models;
using ClosedXML.Excel;
using DotNetDBF;
using Microsoft.Data.Sqlite;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;

namespace ArisPayrollCompanion.Views
{
    public partial class ActiveLoansReportWindow : Window
    {
        private List<ActiveLoanReportItem> _masterList = new();
        private string _dbfFolderPath = "";

        public ActiveLoansReportWindow()
        {
            InitializeComponent();

            _dbfFolderPath = App.DbfFolderPath;
            if (string.IsNullOrEmpty(_dbfFolderPath) || !Directory.Exists(_dbfFolderPath))
            {
                _dbfFolderPath = @"C:\PayrollData";
            }

            LoadReportData();
        }

        private DataTable ReadDbfManually(string fileName)
        {
            DataTable dt = new DataTable();
            string filePath = Path.Combine(_dbfFolderPath, fileName);

            if (!File.Exists(filePath))
            {
                filePath = Path.Combine(_dbfFolderPath, fileName.ToLower());
                if (!File.Exists(filePath)) return dt;
            }

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
                        if (!dt.Columns.Contains(name)) dt.Columns.Add(name);
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

        private bool CodesMatch(string a, string b)
        {
            if (string.IsNullOrWhiteSpace(a) || string.IsNullOrWhiteSpace(b)) return false;
            string digitsA = new string(a.Where(char.IsDigit).ToArray());
            string digitsB = new string(b.Where(char.IsDigit).ToArray());
            if (string.IsNullOrEmpty(digitsA) || string.IsNullOrEmpty(digitsB)) return false;
            if (long.TryParse(digitsA, out long numA) && long.TryParse(digitsB, out long numB))
                return numA == numB;
            return a.Trim() == b.Trim();
        }

        private double ParseAmount(object value)
        {
            if (value == null) return 0;
            string s = value.ToString()?.Trim() ?? "";
            if (string.IsNullOrEmpty(s)) return 0;

            s = s.Replace(",", "");

            return double.TryParse(s, out double result) ? result : 0;
        }

        private void LoadReportData()
        {
            try
            {
                _masterList.Clear();

                List<Employee> activeEmployees = new();
                using (var connection = new SqliteConnection(DatabaseHelper.GetConnectionString()))
                {
                    connection.Open();
                    string query = @"
                        SELECT EmployeeCode, CardNumber, EmployeeName, Department, Category 
                        FROM Employees 
                        WHERE Status = 'Y' OR Status = 'Active' OR Status = 'A' OR Status = '1'";

                    using var cmd = new SqliteCommand(query, connection);
                    using var reader = cmd.ExecuteReader();
                    while (reader.Read())
                    {
                        activeEmployees.Add(new Employee
                        {
                            EmployeeCode = reader["EmployeeCode"].ToString() ?? "",
                            CardNumber = reader["CardNumber"].ToString() ?? "",
                            EmployeeName = reader["EmployeeName"].ToString() ?? "",
                            Department = reader["Department"].ToString() ?? "",
                            Category = reader["Category"].ToString() ?? ""
                        });
                    }
                }

                DataTable monsumDt = ReadDbfManually("monsum.DBF");
                if (monsumDt.Rows.Count == 0)
                {
                    MessageBox.Show("Could not read monsum.DBF data.", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                foreach (var emp in activeEmployees)
                {
                    var rawMatches = new List<(string month, double emiDed, double paidLoan, double loanRec)>();

                    foreach (DataRow row in monsumDt.Rows)
                    {
                        string dbfEmpCode = monsumDt.Columns.Contains("EMPCD") ? row["EMPCD"].ToString() ?? "" : "";

                        bool isMatch = CodesMatch(dbfEmpCode, emp.EmployeeCode);
                        if (!isMatch) continue;

                        string month = monsumDt.Columns.Contains("YRMONTH") ? row["YRMONTH"].ToString()?.Trim() ?? "" : "";
                        if (string.IsNullOrEmpty(month) && monsumDt.Columns.Contains("PAY_MONTH"))
                            month = row["PAY_MONTH"].ToString()?.Trim() ?? "";

                        double emiDed = monsumDt.Columns.Contains("LOAN_DED") ? ParseAmount(row["LOAN_DED"]) : 0;
                        double paidLoan = monsumDt.Columns.Contains("PAID_LOAN") ? ParseAmount(row["PAID_LOAN"]) : 0;
                        double loanRec = monsumDt.Columns.Contains("LOAN_REC") ? ParseAmount(row["LOAN_REC"]) : 0;

                        rawMatches.Add((month, emiDed, paidLoan, loanRec));
                    }

                    var monthlyData = rawMatches
                        .Where(m => !string.IsNullOrEmpty(m.month))
                        .GroupBy(m => m.month)
                        .Select(g => g.Last())
                        .OrderBy(m => m.month)
                        .ToList();

                    var lastRecord = monthlyData.LastOrDefault();

                    double currentTotalApproved = 0;
                    double latestRecovered = 0;
                    double currentMonthEmiDed = 0;

                    if (lastRecord != default)
                    {
                        currentTotalApproved = lastRecord.paidLoan;
                        latestRecovered = lastRecord.loanRec;
                        currentMonthEmiDed = lastRecord.emiDed;
                    }

                    decimal currentOutstanding = (decimal)(currentTotalApproved - latestRecovered - currentMonthEmiDed);
                    if (currentOutstanding < 0) currentOutstanding = 0;

                    if (currentOutstanding > 0)
                    {
                        _masterList.Add(new ActiveLoanReportItem
                        {
                            CardNumber = emp.CardNumber,
                            EmployeeCode = emp.EmployeeCode,
                            EmployeeName = string.IsNullOrEmpty(emp.EmployeeName) ? "Name Missing" : emp.EmployeeName,
                            Department = GetDepartmentName(emp.Department),
                            Category = GetCategoryName(emp.Category),
                            PendingLoanAmount = currentOutstanding
                        });
                    }
                }

                PopulateFilters();
                ApplyFilters();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading active loans report: {ex.Message}", "System Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void PopulateFilters()
        {
            var depts = _masterList.Select(x => x.Department).Where(d => !string.IsNullOrEmpty(d)).Distinct().OrderBy(d => d).ToList();
            depts.Insert(0, "All Departments");
            cmbDepartment.ItemsSource = depts;
            cmbDepartment.SelectedIndex = 0;

            var cats = _masterList.Select(x => x.Category).Where(c => !string.IsNullOrEmpty(c)).Distinct().OrderBy(c => c).ToList();
            cats.Insert(0, "All Categories");
            cmbCategory.ItemsSource = cats;
            cmbCategory.SelectedIndex = 0;
        }

        private void Filter_SelectionChanged(object sender, SelectionChangedEventArgs e) => ApplyFilters();
        private void BtnRefresh_Click(object sender, RoutedEventArgs e) => LoadReportData();

        private void ApplyFilters()
        {
            string selectedDept = cmbDepartment.SelectedItem as string ?? "All Departments";
            string selectedCat = cmbCategory.SelectedItem as string ?? "All Categories";

            var filtered = _masterList.Where(x =>
                (selectedDept == "All Departments" || x.Department == selectedDept) &&
                (selectedCat == "All Categories" || x.Category == selectedCat)
            ).OrderBy(x => int.TryParse(x.EmployeeCode, out int n) ? n : int.MaxValue).ToList();

            dgReport.ItemsSource = filtered;
            txtTotalActive.Text = $"Total Active with Loans: {filtered.Count}";
            txtTotalLoan.Text = $"Total Pending Loans: ₹{filtered.Sum(x => x.PendingLoanAmount):N2}";
        }

        private string GetDepartmentName(string code) => code?.Trim().TrimStart('0') switch { "1" => "OFFICE", "2" => "CASTING", "3" => "MACHINE", "4" => "BUFF & AMRI", "5" => "CHROM", "6" => "FITTING / OTHER", "7" => "OTHER NEW", _ => code };
        private string GetCategoryName(string code) => code?.Trim().TrimStart('0') switch { "1" => "STAFF", "2" => "CASTING_10", "3" => "WORKER", "4" => "MONTHLY_8", "5" => "GENERAL / PIC RATE", "6" => "GENERAL ODD", "7" => "MONTHLY_10", "8" => "TRANSPORT", "10" => "PLUMBER", "11" => "PIC RATE GENERAL", "12" => "WORKER_9", _ => code };

        private void btnExportExcel_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var list = dgReport.ItemsSource as List<ActiveLoanReportItem>;
                if (list == null || list.Count == 0) return;

                SaveFileDialog sfd = new SaveFileDialog { Filter = "Excel Workbook|*.xlsx", FileName = $"Active_Employees_Loans_{DateTime.Now:yyyyMMdd}.xlsx" };
                if (sfd.ShowDialog() == true)
                {
                    using (var wb = new XLWorkbook())
                    {
                        var ws = wb.Worksheets.Add("Active Loans");
                        ws.Cell(1, 1).InsertTable(list);
                        wb.SaveAs(sfd.FileName);
                    }
                    MessageBox.Show("Report exported successfully!", "Export Complete", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Export error: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}