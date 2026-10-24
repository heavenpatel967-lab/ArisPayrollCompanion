using ArisPayrollCompanion.Helpers;
using ClosedXML.Excel;
using Microsoft.Data.Sqlite;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows;

namespace ArisPayrollCompanion.Views
{
    public class ActiveLeaveItem
    {
        public string EmpCode { get; set; } = "";
        public string EmpName { get; set; } = "";
        public string DepartureDate { get; set; } = "";
        public string ExpectedReturn { get; set; } = "";
        public string Reason { get; set; } = "";
        public int LateDays { get; set; }

        public string LateDaysText => LateDays > 0 ? $"{LateDays} Days" : "-";
        public string LateColor => LateDays > 0 ? "#FF5252" : "Black";
    }

    public partial class ActiveLeavesReportWindow : Window
    {
        public ActiveLeavesReportWindow()
        {
            InitializeComponent();
            LoadActiveLeaves();
        }

        private void LoadActiveLeaves()
        {
            try
            {
                List<ActiveLeaveItem> leaveList = new List<ActiveLeaveItem>();
                int totalOverdue = 0;

                using (var connection = new SqliteConnection(DatabaseHelper.GetConnectionString()))
                {
                    connection.Open();

                    // FIXED: Join with 'Employees' table and select 'EmployeeName' instead of 'EmployeeProfiles'
                    string query = @"
                        SELECT l.EmployeeCode, e.EmployeeName, l.DepartureDate, l.ExpectedReturnDate, l.Reason 
                        FROM LongLeaveTracking l
                        LEFT JOIN Employees e ON l.EmployeeCode = e.EmployeeCode
                        WHERE l.ActualReturnDate IS NULL OR l.ActualReturnDate = ''
                        ORDER BY l.DepartureDate DESC";

                    using (var cmd = new SqliteCommand(query, connection))
                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            string code = reader["EmployeeCode"]?.ToString() ?? "";
                            string name = reader["EmployeeName"]?.ToString() ?? "";

                            if (string.IsNullOrWhiteSpace(name)) name = "Unknown Name";

                            string depDate = reader["DepartureDate"]?.ToString() ?? "";
                            string expDate = reader["ExpectedReturnDate"]?.ToString() ?? "";
                            string reason = reader["Reason"]?.ToString() ?? "";

                            int lateDays = 0;
                            // Dynamically calculate if they are currently late!
                            if (DateTime.TryParse(expDate, out DateTime expectedReturnDT))
                            {
                                if (expectedReturnDT < DateTime.Today)
                                {
                                    lateDays = (DateTime.Today - expectedReturnDT).Days;
                                    totalOverdue++;
                                }
                            }

                            leaveList.Add(new ActiveLeaveItem
                            {
                                EmpCode = code,
                                EmpName = name,
                                DepartureDate = depDate,
                                ExpectedReturn = expDate,
                                LateDays = lateDays,
                                Reason = reason
                            });
                        }
                    }
                }

                dgActiveLeaves.ItemsSource = leaveList;
                txtTotalOnLeave.Text = leaveList.Count.ToString();
                txtTotalOverdue.Text = totalOverdue.ToString();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading active leaves: {ex.Message}");
            }
        }

        private void btnExportExcel_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                SaveFileDialog sfd = new SaveFileDialog
                {
                    Filter = "Excel Workbook|*.xlsx",
                    FileName = $"Active_Leaves_Report_{DateTime.Now:yyyyMMdd}.xlsx"
                };

                if (sfd.ShowDialog() == true)
                {
                    using (var workbook = new XLWorkbook())
                    {
                        var ws = workbook.Worksheets.Add("Active Leaves");
                        ws.Cell("A1").Value = "Live Active Leaves Report";
                        ws.Range("A1:E1").Merge().Style.Font.SetBold().Font.FontSize = 14;

                        ws.Cell("A3").Value = "Emp Code";
                        ws.Cell("B3").Value = "Employee Name";
                        ws.Cell("C3").Value = "Departure";
                        ws.Cell("D3").Value = "Expected Return";
                        ws.Cell("E3").Value = "Late Days";
                        ws.Cell("F3").Value = "Reason / Notes";
                        ws.Range("A3:F3").Style.Font.SetBold().Fill.SetBackgroundColor(XLColor.LightGray);

                        int row = 4;
                        if (dgActiveLeaves.ItemsSource is List<ActiveLeaveItem> leaveList)
                        {
                            foreach (var item in leaveList)
                            {
                                ws.Cell(row, 1).Value = item.EmpCode;
                                ws.Cell(row, 2).Value = item.EmpName;
                                ws.Cell(row, 3).Value = item.DepartureDate;
                                ws.Cell(row, 4).Value = item.ExpectedReturn;

                                var lateCell = ws.Cell(row, 5);
                                lateCell.Value = item.LateDaysText;
                                if (item.LateDays > 0)
                                {
                                    lateCell.Style.Font.FontColor = XLColor.Red;
                                    lateCell.Style.Font.SetBold();
                                }

                                ws.Cell(row, 6).Value = item.Reason;
                                row++;
                            }
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
    }
}