using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Data.Sqlite;
using ArisPayrollCompanion.Models;
using ArisPayrollCompanion.Helpers;
using ClosedXML.Excel;
using Microsoft.Win32;

namespace ArisPayrollCompanion.Views
{
    public partial class DailyReportsWindow : Window
    {
        private List<ReportResult> _allLateRecords = new List<ReportResult>();
        private List<ReportResult> _allAbsentRecords = new List<ReportResult>();

        public DailyReportsWindow()
        {
            InitializeComponent();
            dpReportDate.SelectedDate = DateTime.Today;

            // 🔥 LOAD SAVED PATH: Checks if you already set it earlier in the session
            if (App.Current.Properties["MonDataPath"] is string saved && File.Exists(saved))
            {
                txtMonDataPath.Text = saved;
            }
            else
            {
                string defaultPath = Path.Combine(App.DbfFolderPath, "MONDATA.DBF");
                if (!File.Exists(defaultPath)) defaultPath = Path.Combine(App.DbfFolderPath, "mondata.dbf");
                if (File.Exists(defaultPath)) txtMonDataPath.Text = defaultPath;
            }
        }

        private void btnBrowseMonData_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog { Title = "Select MONDATA.DBF", Filter = "DBF Files (*.dbf)|*.dbf", InitialDirectory = App.DbfFolderPath };
            if (ofd.ShowDialog() == true)
            {
                txtMonDataPath.Text = ofd.FileName;
                App.Current.Properties["MonDataPath"] = ofd.FileName; // Save it to global memory
            }
        }

        private void btnRunReport_Click(object sender, RoutedEventArgs e)
        {
            string monDataFile = txtMonDataPath.Text.Trim();
            if (string.IsNullOrWhiteSpace(monDataFile) || !File.Exists(monDataFile))
            {
                MessageBox.Show("Please select a valid MONDATA.DBF file.", "Missing File", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            App.Current.Properties["MonDataPath"] = monDataFile; // Lock in the path

            if (dpReportDate.SelectedDate == null) return;
            DateTime selected = dpReportDate.SelectedDate.Value;
            bool isWednesday = selected.DayOfWeek == DayOfWeek.Wednesday;

            string[] formats = { selected.ToString("yyyyMMdd"), selected.ToString("dd-MM-yyyy"), selected.ToString("MM/dd/yyyy"), selected.ToString("yyyy-MM-dd") };

            var readerService = new ArisPayrollCompanion.Services.DbfReaderService();
            List<DailyAttendance> attendanceRecords = new List<DailyAttendance>();

            foreach (string fmt in formats)
            {
                attendanceRecords = readerService.GetDailyAttendance(monDataFile, fmt);
                if (attendanceRecords.Count > 0) break;
            }

            if (attendanceRecords.Count == 0)
            {
                MessageBox.Show($"No attendance data found for {selected:dd-MMM-yyyy}.", "No Data Found", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var empDb = new Dictionary<string, Employee>();
            try
            {
                using (var conn = new SqliteConnection(DatabaseHelper.GetConnectionString()))
                {
                    conn.Open();
                    // Grab CardNumber as well so we can search by it
                    using (var cmd = new SqliteCommand("SELECT EmployeeCode, EmployeeName, Department, CardNumber FROM EmployeeProfiles", conn))
                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            string code = reader.IsDBNull(0) ? "" : reader.GetString(0);
                            empDb[code] = new Employee
                            {
                                EmployeeName = reader.IsDBNull(1) ? "" : reader.GetString(1),
                                Department = reader.IsDBNull(2) ? "" : reader.GetString(2),
                                CardNumber = reader.IsDBNull(3) ? "" : reader.GetString(3)
                            };
                        }
                    }
                }
            }
            catch { }

            _allLateRecords.Clear();
            _allAbsentRecords.Clear();

            foreach (var record in attendanceRecords)
            {
                empDb.TryGetValue(record.EmployeeCode, out Employee empInfo);
                string name = empInfo?.EmployeeName ?? "Unknown / Not Synced";
                string dept = empInfo?.Department ?? "";
                string card = empInfo?.CardNumber ?? record.EmployeeCode;

                if (record.InTime > 8.32 && !record.IsAbsent)
                {
                    _allLateRecords.Add(new ReportResult { EmployeeCode = record.EmployeeCode, CardNumber = card, EmployeeName = name, Department = dept, PunchTimeStr = record.InTime.ToString("0.00"), Status = "Late" });
                }

                if (record.IsAbsent)
                {
                    string wOff = record.WeekOffDay?.ToUpper().Trim() ?? "";
                    if (isWednesday && wOff.Contains("WED")) continue;

                    _allAbsentRecords.Add(new ReportResult { EmployeeCode = record.EmployeeCode, CardNumber = card, EmployeeName = name, Department = dept, PunchTimeStr = "-", Status = "Absent" });
                }
            }

            ApplySearchFilter(); // Send it to the grids!
        }

        // 🔥 NEW: Live Search Feature
        private void txtSearch_TextChanged(object sender, TextChangedEventArgs e)
        {
            ApplySearchFilter();
        }

        private void ApplySearchFilter()
        {
            if (_allLateRecords == null || _allAbsentRecords == null) return;
            string search = txtSearch.Text.ToLower().Trim();

            if (string.IsNullOrEmpty(search))
            {
                dgLate.ItemsSource = _allLateRecords.OrderByDescending(x => double.Parse(x.PunchTimeStr)).ToList();
                dgAbsent.ItemsSource = _allAbsentRecords.OrderBy(x => x.CardNumber).ToList();
            }
            else
            {
                dgLate.ItemsSource = _allLateRecords.Where(x => x.CardNumber.Contains(search) || x.EmployeeCode.Contains(search) || x.EmployeeName.ToLower().Contains(search))
                                                    .OrderByDescending(x => double.Parse(x.PunchTimeStr)).ToList();

                dgAbsent.ItemsSource = _allAbsentRecords.Where(x => x.CardNumber.Contains(search) || x.EmployeeCode.Contains(search) || x.EmployeeName.ToLower().Contains(search))
                                                        .OrderBy(x => x.CardNumber).ToList();
            }
        }

        private void btnExportExcel_Click(object sender, RoutedEventArgs e)
        {
            bool isLateTab = tabControl.SelectedIndex == 0;
            var currentData = isLateTab ? dgLate.ItemsSource as List<ReportResult> : dgAbsent.ItemsSource as List<ReportResult>;
            string reportName = isLateTab ? "LatePunching" : "Absentees";

            if (currentData == null || currentData.Count == 0) return;

            SaveFileDialog sfd = new SaveFileDialog { Filter = "Excel Workbook|*.xlsx", FileName = $"{reportName}_{dpReportDate.SelectedDate.Value:dd-MMM-yyyy}.xlsx" };

            if (sfd.ShowDialog() == true)
            {
                try
                {
                    using (var wb = new XLWorkbook())
                    {
                        var ws = wb.Worksheets.Add(reportName);
                        // Clean export to remove backend code columns
                        var exportData = currentData.Select(x => new { Card_No = x.CardNumber, Code = x.EmployeeCode, Name = x.EmployeeName, Department = x.Department, Value = isLateTab ? x.PunchTimeStr : x.Status });
                        ws.Cell(1, 1).InsertTable(exportData);
                        ws.Columns().AdjustToContents();
                        wb.SaveAs(sfd.FileName);
                    }
                    MessageBox.Show("Excel report generated successfully!", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex) { MessageBox.Show(ex.Message, "Export Error", MessageBoxButton.OK, MessageBoxImage.Error); }
            }
        }
    }
}