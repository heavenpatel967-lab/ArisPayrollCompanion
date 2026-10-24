using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using Microsoft.Win32;
using ClosedXML.Excel;
using ArisPayrollCompanion.Models;

namespace ArisPayrollCompanion.Views
{
    // Models for the Unified UI Bindings
    public class LedgerItem
    {
        public string MonthYear { get; set; }
        public string LoanAmount { get; set; }
        public string EmiAmount { get; set; }
        public string EmiColor { get; set; }
        public string EmiWeight { get; set; }
    }

    public class HealthBlock
    {
        public string Color { get; set; }
        public string Month { get; set; }
    }

    public partial class LoanReportWindow : Window
    {
        private readonly Employee _employee;

        // Dynamically references the globally saved path
        private readonly string _dbfFolderPath = App.DbfFolderPath;

        public LoanReportWindow(Employee employee)
        {
            InitializeComponent();
            _employee = employee;
            LoadLoanDataFromSalaryRegisters();
        }

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
            string s = value.ToString().Trim();
            if (string.IsNullOrEmpty(s)) return 0;
            return double.TryParse(s, out double result) ? result : 0;
        }

        private void LoadLoanDataFromSalaryRegisters()
        {
            DataTable dt = ReadDbfManually("monsum.DBF");
            string targetCode = _employee.EmployeeCode;
            string targetCard = _employee.CardNumber;

            var rawMatches = new List<(string month, double emiDed, double paidLoan, double loanRec)>();

            foreach (DataRow row in dt.Rows)
            {
                string dbfEmpCode = dt.Columns.Contains("EMPCD") ? row["EMPCD"].ToString() : "";
                string dbfCardNo = dt.Columns.Contains("CARDNO") ? row["CARDNO"].ToString() : "";

                bool isMatch = CodesMatch(dbfEmpCode, targetCode) ||
                               (!string.IsNullOrWhiteSpace(targetCard) && CodesMatch(dbfCardNo, targetCard));
                if (!isMatch) continue;

                string month = dt.Columns.Contains("YRMONTH") ? row["YRMONTH"].ToString().Trim() : "";
                if (string.IsNullOrEmpty(month) && dt.Columns.Contains("PAY_MONTH"))
                    month = row["PAY_MONTH"].ToString().Trim();

                double emiDed = dt.Columns.Contains("LOAN_DED") ? ParseAmount(row["LOAN_DED"]) : 0;
                double paidLoan = dt.Columns.Contains("PAID_LOAN") ? ParseAmount(row["PAID_LOAN"]) : 0;
                double loanRec = dt.Columns.Contains("LOAN_REC") ? ParseAmount(row["LOAN_REC"]) : 0;

                rawMatches.Add((month, emiDed, paidLoan, loanRec));
            }

            var monthlyData = rawMatches
                .Where(m => !string.IsNullOrEmpty(m.month))
                .GroupBy(m => m.month)
                .Select(g => g.Last())
                .OrderBy(m => m.month)
                .ToList();

            List<LedgerItem> unifiedList = new List<LedgerItem>();
            double totalDed = 0;
            int monthsTotal = 0;
            int monthsPaid = 0;

            foreach (var m in monthlyData)
            {
                bool hasPending = m.paidLoan > 0 || m.loanRec > 0;
                bool paidEmi = m.emiDed > 0;

                if (hasPending || paidEmi)
                {
                    monthsTotal++;
                    if (paidEmi)
                    {
                        monthsPaid++;
                        totalDed += m.emiDed;
                    }

                    double pendingAmount = m.paidLoan - m.loanRec;

                    unifiedList.Add(new LedgerItem
                    {
                        MonthYear = m.month,
                        LoanAmount = hasPending ? $"₹ {pendingAmount:N2}" : "-",
                        EmiAmount = paidEmi ? $"₹ {m.emiDed:N2}" : "-",
                        EmiColor = paidEmi ? "#FF5252" : "#333333", // Red if paid, dark grey if missing
                        EmiWeight = paidEmi ? "Bold" : "Normal"
                    });
                }
            }

            // Generate Repayment Health (Last 6 Months)
            var last6 = monthlyData.Where(m => m.paidLoan > 0 || m.loanRec > 0 || m.emiDed > 0).TakeLast(6).ToList();
            List<HealthBlock> healthBlocks = new List<HealthBlock>();
            foreach (var m in last6)
            {
                healthBlocks.Add(new HealthBlock
                {
                    Month = m.month,
                    Color = (m.emiDed > 0) ? "#27AE60" : "#E74C3C" // Green if paid, Red if missed
                });
            }

            double latestTotalLoan = 0;
            double latestRecovered = 0;

            var lastValidRecord = monthlyData.LastOrDefault(m => m.paidLoan > 0 || m.loanRec > 0);
            if (lastValidRecord != default)
            {
                latestTotalLoan = lastValidRecord.paidLoan;
                latestRecovered = lastValidRecord.loanRec;
            }

            double currentOutstanding = latestTotalLoan - latestRecovered;

            // Bind UI Elements
            dgUnifiedLedger.ItemsSource = unifiedList;
            icLast6Months.ItemsSource = healthBlocks;

            txtRepaymentFraction.Text = $"{monthsPaid} / {monthsTotal} Paid";
            txtTopTotalApproved.Text = $"₹ {latestTotalLoan:N2}";
            txtTopTotalEMI.Text = $"₹ {latestRecovered:N2}";
            txtTopOutstanding.Text = $"₹ {currentOutstanding:N2}";

            txtBottomTotalGiven.Text = $"₹ {currentOutstanding:N2}";
            txtBottomTotalEMI.Text = $"₹ {totalDed:N2}";
        }

        private void btnExportPDF_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Controls.PrintDialog printDlg = new System.Windows.Controls.PrintDialog();
            if (printDlg.ShowDialog() == true)
            {
                ActionPanel.Visibility = Visibility.Hidden; // Hides buttons cleanly
                printDlg.PrintVisual(this.Content as UIElement, $"Loan Ledger - {_employee.EmployeeCode}");
                ActionPanel.Visibility = Visibility.Visible;
            }
        }

        private void btnExportExcel_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // Explicitly use the Microsoft.Win32 version to prevent ambiguity
                Microsoft.Win32.SaveFileDialog sfd = new Microsoft.Win32.SaveFileDialog
                {
                    Filter = "Excel Workbook|*.xlsx",
                    FileName = $"Unified_LoanLedger_{_employee.EmployeeCode}.xlsx"
                };

                if (sfd.ShowDialog() == true)
                {
                    using (var workbook = new XLWorkbook())
                    {
                        var ws = workbook.Worksheets.Add("Loan Report");

                        // Headers & Top Summary
                        ws.Cell("A1").Value = $"Unified Loan Ledger - {_employee.EmployeeCode}";
                        ws.Range("A1:D1").Merge().Style.Font.SetBold().Font.FontSize = 14;

                        ws.Cell("A3").Value = "Total Approved:";
                        ws.Cell("B3").Value = txtTopTotalApproved.Text;

                        ws.Cell("C3").Value = "Total EMI Deducted:";
                        ws.Cell("D3").Value = txtTopTotalEMI.Text;

                        ws.Cell("A4").Value = "Current Outstanding:";
                        ws.Cell("B4").Value = txtTopOutstanding.Text;
                        ws.Range("A3:D4").Style.Font.SetBold();

                        // Build Unified Grid
                        ws.Cell("A6").Value = "Month/Year";
                        ws.Cell("B6").Value = "Pending Loan Amount";
                        ws.Cell("C6").Value = "EMI Deducted";
                        ws.Range("A6:C6").Style.Font.SetBold().Fill.SetBackgroundColor(XLColor.LightGray);

                        int row = 7;
                        if (dgUnifiedLedger.ItemsSource is List<LedgerItem> ledgerList)
                        {
                            foreach (var item in ledgerList)
                            {
                                ws.Cell(row, 1).Value = item.MonthYear;
                                ws.Cell(row, 2).Value = item.LoanAmount;
                                ws.Cell(row, 3).Value = item.EmiAmount;

                                if (item.EmiAmount != "-")
                                {
                                    ws.Cell(row, 3).Style.Font.FontColor = XLColor.Red;
                                    ws.Cell(row, 3).Style.Font.SetBold();
                                }
                                row++;
                            }
                        }

                        ws.Columns().AdjustToContents();
                        workbook.SaveAs(sfd.FileName);

                        // Explicitly use the System.Windows version of MessageBox
                        System.Windows.MessageBox.Show("Excel report generated successfully!", "Success", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Information);
                    }
                }
            }
            catch (System.Exception ex)
            {
                // Explicitly use the System.Windows version of MessageBox
                System.Windows.MessageBox.Show($"Error exporting to Excel: {ex.Message}", "Export Error", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
            }
        }
    }
}