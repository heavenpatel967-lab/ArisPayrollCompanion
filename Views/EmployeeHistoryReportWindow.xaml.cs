using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using Microsoft.Data.Sqlite;
using ArisPayrollCompanion.Models;
using ArisPayrollCompanion.Helpers;
using ClosedXML.Excel;
using Microsoft.Win32;

namespace ArisPayrollCompanion.Views
{
    public class EmployeeDropdownItem
    {
        public string EmployeeCode { get; set; }
        public string EmployeeName { get; set; }
        public string DisplayName => $"{EmployeeCode} - {EmployeeName}";
    }

    public partial class EmployeeHistoryReportWindow : Window
    {
        public EmployeeHistoryReportWindow()
        {
            InitializeComponent();

            dpStart.SelectedDate = new DateTime(DateTime.Today.Year, DateTime.Today.Month, 1);
            dpEnd.SelectedDate = DateTime.Today;

            // 🔥 LOAD SAVED PATH: Shares the exact same memory as the Daily Reports Window
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

            LoadEmployeesIntoDropdown();
        }

        private void LoadEmployeesIntoDropdown()
        {
            var empList = new List<EmployeeDropdownItem>();
            try
            {
                using (var conn = new SqliteConnection(DatabaseHelper.GetConnectionString()))
                {
                    conn.Open();
                    using (var cmd = new SqliteCommand("SELECT EmployeeCode, EmployeeName FROM EmployeeProfiles ORDER BY EmployeeCode", conn))
                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            empList.Add(new EmployeeDropdownItem
                            {
                                EmployeeCode = reader.IsDBNull(0) ? "" : reader.GetString(0),
                                EmployeeName = reader.IsDBNull(1) ? "" : reader.GetString(1)
                            });
                        }
                    }
                }
                cmbEmployee.ItemsSource = empList;
                cmbEmployee.DisplayMemberPath = "DisplayName";
                cmbEmployee.SelectedValuePath = "EmployeeCode";
            }
            catch { }
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

            if (dpStart.SelectedDate == null || dpEnd.SelectedDate == null) return;

            string targetEmp = cmbEmployee.SelectedValue?.ToString();
            if (string.IsNullOrEmpty(targetEmp) && !string.IsNullOrWhiteSpace(cmbEmployee.Text))
            {
                targetEmp = cmbEmployee.Text.Split('-')[0].Trim();
            }

            if (string.IsNullOrWhiteSpace(targetEmp)) return;

            var readerService = new ArisPayrollCompanion.Services.DbfReaderService();
            var history = readerService.GetEmployeeAttendanceHistory(monDataFile, targetEmp, dpStart.SelectedDate.Value, dpEnd.SelectedDate.Value);

            if (history.Count == 0)
            {
                MessageBox.Show($"No attendance data found for Employee {targetEmp} in this date range.", "No Data", MessageBoxButton.OK, MessageBoxImage.Information);
            }

            dgHistory.ItemsSource = history;
        }

        private void btnExportExcel_Click(object sender, RoutedEventArgs e)
        {
            var data = dgHistory.ItemsSource as List<EmployeeDailyPunch>;
            if (data == null || data.Count == 0) return;

            SaveFileDialog sfd = new SaveFileDialog
            {
                Filter = "Excel Workbook|*.xlsx",
                FileName = $"History_{cmbEmployee.Text.Split('-')[0].Trim()}_{dpStart.SelectedDate.Value:ddMMM}_to_{dpEnd.SelectedDate.Value:ddMMM}.xlsx"
            };

            if (sfd.ShowDialog() == true)
            {
                try
                {
                    using (var wb = new XLWorkbook())
                    {
                        var ws = wb.Worksheets.Add("History Report");
                        var exportData = data.Select(x => new { Date = x.DateStr, In_Time = x.InTime, Out_Time = x.OutTime, Status = x.Status });
                        ws.Cell(1, 1).InsertTable(exportData);
                        ws.Columns().AdjustToContents();
                        wb.SaveAs(sfd.FileName);
                    }
                    MessageBox.Show("Report exported successfully!", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex) { MessageBox.Show(ex.Message, "Export Error", MessageBoxButton.OK, MessageBoxImage.Error); }
            }
        }
    }
}