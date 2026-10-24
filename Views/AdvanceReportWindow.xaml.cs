using ArisPayrollCompanion.Models;
using ClosedXML.Excel;
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
    // Added to support the Last 6 Months colored dots in your XAML
    public class AdvanceTrendItem
    {
        public string Month { get; set; }
        public string Color { get; set; }
    }

    public class AdvanceLedgerItem
    {
        public string MonthYear { get; set; }
        public double RawAmount { get; set; }
        public string StatusColor { get; set; }

        // EXACT MATCH: This connects directly to Text="{Binding AmountDisplay}" in your XAML
        public string AmountDisplay => $"₹ {RawAmount:N2}";
    }

    public partial class AdvanceReportWindow : Window
    {
        private readonly Employee _employee;
        private readonly string _dbfFolderPath = App.DbfFolderPath;
        private List<AdvanceLedgerItem> _allAdvances = new List<AdvanceLedgerItem>();

        private readonly string _limitConfigFile;
        private bool _isLoaded = false;

        public AdvanceReportWindow(Employee employee)
        {
            InitializeComponent();
            _employee = employee;

            // Unique file limit per employee
            _limitConfigFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, $"advance_limit_{_employee.EmployeeCode}.txt");

            if (File.Exists(_limitConfigFile))
            {
                string savedLimit = File.ReadAllText(_limitConfigFile).Trim();
                if (!string.IsNullOrEmpty(savedLimit))
                {
                    txtAdvanceLimit.Text = savedLimit;
                }
            }
            else
            {
                txtAdvanceLimit.Text = "5000";
            }

            _isLoaded = true;

            LoadAdvanceDataFromSalaryRegisters();
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
                        {
                            row[field.Item1] = Encoding.Latin1.GetString(rowData, offset, field.Item2).Trim();
                        }
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
            s = s.Replace(",", "");
            return double.TryParse(s, out double result) ? result : 0;
        }

        private void LoadAdvanceDataFromSalaryRegisters()
        {
            DataTable dt = ReadDbfManually("monsum.DBF");
            string targetCode = _employee.EmployeeCode;
            string targetCard = _employee.CardNumber;

            var rawMatches = new List<(string month, double amt)>();

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

                double amt = dt.Columns.Contains("ADVANCE") ? ParseAmount(row["ADVANCE"]) : 0;

                rawMatches.Add((month, amt));
            }

            var distinctMatches = rawMatches.Distinct().OrderBy(r => r.month).ToList();

            _allAdvances.Clear();
            double totalAdvance = 0;

            foreach (var m in distinctMatches)
            {
                if (m.amt != 0)
                {
                    totalAdvance += m.amt;
                    string label = string.IsNullOrEmpty(m.month) ? "Advance Taken" : m.month;
                    _allAdvances.Add(new AdvanceLedgerItem
                    {
                        MonthYear = label,
                        RawAmount = m.amt,
                        StatusColor = "#333333"
                    });
                }
            }

            txtTopTotalAdvance.Text = $"₹ {totalAdvance:N2}";
            txtBottomTotalAdvance.Text = $"₹ {totalAdvance:N2}";

            ApplyLimitAndRefreshGrid();
        }

        private void ApplyLimitAndRefreshGrid()
        {
            if (_allAdvances == null || !_allAdvances.Any()) return;

            double limit = 0;
            if (txtAdvanceLimit != null && double.TryParse(txtAdvanceLimit.Text, out double parsed))
            {
                limit = parsed;
            }

            int safeCount = 0;
            foreach (var item in _allAdvances)
            {
                if (limit > 0 && item.RawAmount > limit)
                {
                    item.StatusColor = "#E74C3C"; // Red (High)
                }
                else
                {
                    item.StatusColor = "#27AE60"; // Green (Safe)
                    safeCount++;
                }
            }

            if (dgAdvanceHistory.ItemsSource == null)
            {
                dgAdvanceHistory.ItemsSource = _allAdvances;
            }
            else
            {
                dgAdvanceHistory.Items.Refresh();
            }

            // EXACT MATCH: Updates your specific txtHealthFraction TextBlock
            if (txtHealthFraction != null)
            {
                txtHealthFraction.Text = $"{safeCount} / {_allAdvances.Count} Safe";
            }

            // EXACT MATCH: Connects the colored dots to your icLast6Months ItemsControl
            if (icLast6Months != null)
            {
                var last6 = _allAdvances.Skip(Math.Max(0, _allAdvances.Count - 6))
                    .Select(a => new AdvanceTrendItem { Month = a.MonthYear, Color = a.StatusColor })
                    .ToList();

                icLast6Months.ItemsSource = last6;
            }
        }

        private void txtAdvanceLimit_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (!_isLoaded) return;

            if (txtAdvanceLimit != null && !string.IsNullOrWhiteSpace(txtAdvanceLimit.Text))
            {
                try
                {
                    File.WriteAllText(_limitConfigFile, txtAdvanceLimit.Text.Trim());
                }
                catch { }
            }

            ApplyLimitAndRefreshGrid();
        }

        private void btnExportExcel_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                SaveFileDialog sfd = new SaveFileDialog
                {
                    Filter = "Excel Workbook|*.xlsx",
                    FileName = $"AdvanceLedger_{_employee.EmployeeCode}.xlsx"
                };

                if (sfd.ShowDialog() == true)
                {
                    using (var workbook = new XLWorkbook())
                    {
                        var ws = workbook.Worksheets.Add("Advance Report");

                        ws.Cell("A1").Value = $"Advance Ledger Report - Employee: {_employee.EmployeeCode}";
                        ws.Range("A1:C1").Merge().Style.Font.SetBold().Font.FontSize = 14;

                        ws.Cell("A3").Value = "Total Advances Taken:";
                        ws.Cell("B3").Value = txtTopTotalAdvance.Text;
                        ws.Range("A3:B3").Style.Font.SetBold();

                        ws.Cell("A5").Value = "Status";
                        ws.Cell("B5").Value = "Month/Year";
                        ws.Cell("C5").Value = "Advance Amount";
                        ws.Range("A5:C5").Style.Font.SetBold().Fill.SetBackgroundColor(XLColor.LightGray);

                        int row = 6;
                        foreach (var item in _allAdvances)
                        {
                            bool isSafe = item.StatusColor == "#27AE60";

                            ws.Cell(row, 1).Value = isSafe ? "Safe" : "High";
                            ws.Cell(row, 2).Value = item.MonthYear;
                            ws.Cell(row, 3).Value = item.RawAmount;

                            if (!isSafe)
                            {
                                ws.Cell(row, 1).Style.Font.FontColor = XLColor.Red;
                                ws.Cell(row, 3).Style.Font.FontColor = XLColor.Red;
                            }
                            row++;
                        }

                        ws.Columns().AdjustToContents();
                        workbook.SaveAs(sfd.FileName);
                        MessageBox.Show("Excel report generated successfully!", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error exporting to Excel: {ex.Message}", "Export Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void btnExportPDF_Click(object sender, RoutedEventArgs e)
        {
            PrintDialog printDlg = new PrintDialog();
            if (printDlg.ShowDialog() == true)
            {
                printDlg.PrintVisual(this.Content as UIElement, $"Advance Ledger - {_employee.EmployeeCode}");
            }
        }
    }
}