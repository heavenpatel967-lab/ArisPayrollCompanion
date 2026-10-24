using ArisPayrollCompanion.Helpers;
using ArisPayrollCompanion.Models;
using Microsoft.Data.Sqlite;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media.Imaging;

namespace ArisPayrollCompanion.Views
{
    public class SummaryHealthItem
    {
        public string Month { get; set; } = "";
        public string Color { get; set; } = "";
    }

    public partial class EmployeeSummaryWindow : Window
    {
        private readonly Employee _employee;
        private readonly string _dbfFolderPath;

        // Removed the hardcoded path. It is now set dynamically in the constructor.
        private readonly string _limitConfigFile;

        public EmployeeSummaryWindow(Employee employee)
        {
            InitializeComponent();
            _employee = employee;

            _dbfFolderPath = App.DbfFolderPath;
            if (string.IsNullOrEmpty(_dbfFolderPath) || !Directory.Exists(_dbfFolderPath))
            {
                _dbfFolderPath = @"C:\PayrollData";
            }

            // Set the unique path to match the Advance window
            _limitConfigFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, $"advance_limit_{_employee.EmployeeCode}.txt");

            LoadDbfData();
            LoadSqliteData();
            LoadLeaveHistory();
            LoadFinancialData();
        }

        private void LoadDbfData()
        {
            txtFullName.Text = _employee.EmployeeName;
            txtCodes.Text = $"{_employee.EmployeeCode} / {_employee.CardNumber}";

            string deptName = GetDepartmentName(_employee.Department);
            txtDept.Text = $"{deptName} - {_employee.Designation}".Trim(' ', '-');

            string rawStatus = _employee.Status?.Trim().ToUpper() ?? "";
            if (rawStatus == "Y" || rawStatus == "ACTIVE" || rawStatus == "A" || rawStatus == "1")
            {
                txtStatus.Text = "Active";
            }
            else
            {
                txtStatus.Text = "Inactive / Left";
            }
        }

        private void LoadSqliteData()
        {
            try
            {
                using (var connection = new SqliteConnection(DatabaseHelper.GetConnectionString()))
                {
                    connection.Open();
                    string query = "SELECT * FROM EmployeeProfiles WHERE EmployeeCode = @Code";

                    using (var command = new SqliteCommand(query, connection))
                    {
                        command.Parameters.AddWithValue("@Code", _employee.EmployeeCode);

                        using (var reader = command.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                string origName = reader["OriginalName"].ToString() ?? "";
                                if (!string.IsNullOrWhiteSpace(origName))
                                {
                                    txtFullName.Text = origName;
                                }

                                txtPhone.Text = reader["PhoneNumber"].ToString() ?? "";
                                txtAadhar.Text = reader["AadharNumber"].ToString() ?? "";

                                string nativePlace = reader["NativePlace"].ToString() ?? "";
                                string village = reader["Village"].ToString() ?? "";
                                txtVillage.Text = string.IsNullOrWhiteSpace(village) ? nativePlace : $"{village}, {nativePlace}";

                                string savedPhotoPath = reader["PhotoFilePath"].ToString() ?? "";
                                if (!string.IsNullOrEmpty(savedPhotoPath) && File.Exists(savedPhotoPath))
                                {
                                    BitmapImage bitmap = new BitmapImage();
                                    bitmap.BeginInit();
                                    bitmap.UriSource = new Uri(savedPhotoPath);
                                    bitmap.CacheOption = BitmapCacheOption.OnLoad;
                                    bitmap.EndInit();

                                    imgProfilePhoto.Source = bitmap;
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading profile data: {ex.Message}");
            }
        }

        private void LoadLeaveHistory()
        {
            try
            {
                using (var connection = new SqliteConnection(DatabaseHelper.GetConnectionString()))
                {
                    connection.Open();
                    string query = "SELECT DepartureDate, ExpectedReturnDate, ActualReturnDate, LateDays, Reason FROM LongLeaveTracking WHERE EmployeeCode = @Code ORDER BY LeaveID ASC";

                    using (var command = new SqliteCommand(query, connection))
                    {
                        command.Parameters.AddWithValue("@Code", _employee.EmployeeCode);

                        DataTable dt = new DataTable();
                        using (var reader = command.ExecuteReader())
                        {
                            dt.Load(reader);
                        }

                        DataTable gridDt = dt.Clone();
                        var newestLeaves = dt.AsEnumerable().Reverse().Take(5);
                        if (newestLeaves.Any())
                        {
                            gridDt = newestLeaves.CopyToDataTable();
                        }
                        dgLeaves.ItemsSource = gridDt.DefaultView;

                        var last6Leaves = dt.AsEnumerable().Reverse().Take(6).Reverse();
                        List<SummaryHealthItem> leaveTimeline = new List<SummaryHealthItem>();
                        int onTimeCount = 0;
                        int totalCount = 0;

                        foreach (DataRow row in last6Leaves)
                        {
                            totalCount++;
                            int lateDays = 0;

                            if (row["LateDays"] != DBNull.Value && int.TryParse(row["LateDays"].ToString(), out int parsedLate))
                            {
                                lateDays = parsedLate;
                            }

                            string actualReturn = row["ActualReturnDate"].ToString() ?? "";
                            string expDateStr = row["ExpectedReturnDate"].ToString() ?? "";

                            if (string.IsNullOrWhiteSpace(actualReturn) && DateTime.TryParse(expDateStr, out DateTime expectedReturnDT))
                            {
                                if (expectedReturnDT < DateTime.Today)
                                {
                                    lateDays = (DateTime.Today - expectedReturnDT).Days;
                                }
                            }

                            string dotColor = "#27AE60";
                            if (lateDays > 7)
                                dotColor = "#FF5252";
                            else if (lateDays > 0)
                                dotColor = "#F1C40F";

                            if (lateDays <= 0) onTimeCount++;

                            leaveTimeline.Add(new SummaryHealthItem
                            {
                                Month = $"Departed: {row["DepartureDate"]}",
                                Color = dotColor
                            });
                        }

                        icLeaveLast6.ItemsSource = leaveTimeline;

                        if (totalCount > 0)
                        {
                            txtLeaveHealth.Text = $"{onTimeCount} / {totalCount} On Time";
                        }
                        else
                        {
                            txtLeaveHealth.Text = "No Leaves Taken";
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading leave history: {ex.Message}");
            }
        }

        private void LoadFinancialData()
        {
            try
            {
                DataTable dt = ReadDbfManually("monsum.DBF");
                if (dt == null || dt.Rows.Count == 0)
                {
                    txtTotalLoan.Text = "₹ 0.00";
                    txtTotalAdvance.Text = "₹ 0.00";
                    return;
                }

                string targetCode = _employee.EmployeeCode;

                var rawLoanMatches = new List<(string month, double emiDed, double paidLoan, double loanRec)>();
                var rawAdvMatches = new List<(string month, double amt)>();

                foreach (DataRow row in dt.Rows)
                {
                    string dbfEmpCode = dt.Columns.Contains("EMPCD") ? row["EMPCD"].ToString() ?? "" : "";

                    bool isMatch = CodesMatch(dbfEmpCode, targetCode);
                    if (!isMatch) continue;

                    string month = dt.Columns.Contains("YRMONTH") ? row["YRMONTH"].ToString()?.Trim() ?? "" : "";
                    if (string.IsNullOrEmpty(month) && dt.Columns.Contains("PAY_MONTH"))
                        month = row["PAY_MONTH"].ToString()?.Trim() ?? "";

                    double emiDed = dt.Columns.Contains("LOAN_DED") ? ParseAmount(row["LOAN_DED"]) : 0;
                    double paidLoan = dt.Columns.Contains("PAID_LOAN") ? ParseAmount(row["PAID_LOAN"]) : 0;
                    double loanRec = dt.Columns.Contains("LOAN_REC") ? ParseAmount(row["LOAN_REC"]) : 0;
                    rawLoanMatches.Add((month, emiDed, paidLoan, loanRec));

                    double advAmt = dt.Columns.Contains("ADVANCE") ? ParseAmount(row["ADVANCE"]) : 0;
                    if (advAmt > 0)
                    {
                        rawAdvMatches.Add((month, advAmt));
                    }
                }

                var monthlyLoanData = rawLoanMatches
                    .Where(m => !string.IsNullOrEmpty(m.month))
                    .GroupBy(m => m.month)
                    .Select(g => g.Last())
                    .OrderBy(m => m.month)
                    .ToList();

                double currentOutstanding = 0;
                var lastValidRecord = monthlyLoanData.LastOrDefault(m => m.paidLoan > 0 || m.loanRec > 0);
                if (lastValidRecord != default)
                {
                    currentOutstanding = lastValidRecord.paidLoan - lastValidRecord.loanRec;
                }

                txtTotalLoan.Text = $"₹ {Math.Max(0, currentOutstanding):N2}";

                int loanTotalMonths = 0;
                int loanPaidMonths = 0;
                List<SummaryHealthItem> loanTimeline = new List<SummaryHealthItem>();

                foreach (var m in monthlyLoanData)
                {
                    double pending = m.paidLoan - m.loanRec;
                    if (pending < 0) pending = 0;

                    if (pending > 0 || m.emiDed > 0)
                    {
                        loanTotalMonths++;
                        if (m.emiDed > 0) loanPaidMonths++;

                        loanTimeline.Add(new SummaryHealthItem
                        {
                            Month = m.month,
                            Color = m.emiDed > 0 ? "#27AE60" : "#FF5252"
                        });
                    }
                }

                if (loanTotalMonths > 0)
                {
                    txtLoanHealthFraction.Text = $"{loanPaidMonths} / {loanTotalMonths} Paid";
                    icLoanLast6Months.ItemsSource = loanTimeline.Skip(Math.Max(0, loanTimeline.Count - 6)).ToList();
                }
                else
                {
                    txtLoanHealthFraction.Text = "No Active Loan";
                }

                var distinctAdvMatches = rawAdvMatches
                    .Where(m => !string.IsNullOrEmpty(m.month))
                    .GroupBy(m => m.month)
                    .Select(g => g.Last())
                    .OrderBy(r => r.month)
                    .ToList();

                double totalAdvance = distinctAdvMatches.Sum(m => m.amt);
                txtTotalAdvance.Text = $"₹ {Math.Max(0, totalAdvance):N2}";

                double advanceLimit = 5000; // The default limit
                if (File.Exists(_limitConfigFile))
                {
                    double.TryParse(File.ReadAllText(_limitConfigFile), out advanceLimit);
                }

                int advTotalMonths = distinctAdvMatches.Count;
                int advSafeMonths = 0;
                List<SummaryHealthItem> advTimeline = new List<SummaryHealthItem>();

                foreach (var m in distinctAdvMatches)
                {
                    bool isSafe = m.amt <= advanceLimit;
                    if (isSafe) advSafeMonths++;

                    advTimeline.Add(new SummaryHealthItem
                    {
                        Month = m.month,
                        Color = isSafe ? "#27AE60" : "#FF5252"
                    });
                }

                if (advTotalMonths > 0)
                {
                    txtAdvanceHealthFraction.Text = $"{advSafeMonths} / {advTotalMonths} Safe";
                    icAdvanceLast6Months.ItemsSource = advTimeline.Skip(Math.Max(0, advTimeline.Count - 6)).ToList();
                }
                else
                {
                    txtAdvanceHealthFraction.Text = "No Advances Taken";
                }
            }
            catch (Exception)
            {
                txtTotalLoan.Text = "Data Error";
                txtTotalAdvance.Text = "Data Error";
            }
        }

        private string GetDepartmentName(string code)
        {
            if (string.IsNullOrWhiteSpace(code)) return "";

            switch (code.Trim().TrimStart('0'))
            {
                case "1": return "OFFICE";
                case "2": return "CASTING";
                case "3": return "MACHINE";
                case "4": return "BUFF & AMRI";
                case "5": return "CHROM";
                case "6": return "FITTING / OTHER";
                case "7": return "OTHER NEW";
                default: return code;
            }
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
            return double.TryParse(s, out double result) ? result : 0;
        }

        private void btnPrint_Click(object sender, RoutedEventArgs e)
        {
            PrintDialog printDlg = new PrintDialog();

            if (printDlg.ShowDialog() == true)
            {
                printDlg.PrintVisual(PrintArea, $"Employee Summary - {_employee.EmployeeName}");
            }
        }

        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}