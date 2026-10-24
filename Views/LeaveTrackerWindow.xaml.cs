using System;
using System.Collections.Generic;
using System.Data;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.IO;
using Microsoft.Win32;
using System.Globalization;
using Microsoft.Data.Sqlite;
using ArisPayrollCompanion.Helpers;

namespace ArisPayrollCompanion.Views
{
    public class LeaveHistoryItem
    {
        public string LeaveID { get; set; }
        public string DepartureDate { get; set; }
        public string ExpectedReturnDate { get; set; }
        public string ActualReturnDate { get; set; }
        public string LateDaysText { get; set; }
        public string Reason { get; set; }
        public string StatusColor { get; set; }
    }

    public partial class LeaveTrackerWindow : Window
    {
        private readonly string _employeeCode;
        private string _editingLeaveId = null;

        public LeaveTrackerWindow(string employeeCode, string employeeName)
        {
            InitializeComponent();
            _employeeCode = employeeCode;

            if (string.IsNullOrEmpty(employeeCode))
                txtHeader.Text = "Active Leaves Tracking (All Employees)";
            else
                txtHeader.Text = $"Leave Tracker: {employeeName} ({employeeCode})";

            LoadLeaveHistory();
        }

        private void LoadLeaveHistory()
        {
            try
            {
                using (var connection = new SqliteConnection(DatabaseHelper.GetConnectionString()))
                {
                    connection.Open();

                    string query = string.IsNullOrEmpty(_employeeCode)
                        ? "SELECT LeaveID, DepartureDate, ExpectedReturnDate, ActualReturnDate, LateDays, Reason FROM LongLeaveTracking WHERE ActualReturnDate IS NULL OR ActualReturnDate = '' ORDER BY ExpectedReturnDate ASC"
                        : "SELECT LeaveID, DepartureDate, ExpectedReturnDate, ActualReturnDate, LateDays, Reason FROM LongLeaveTracking WHERE EmployeeCode = @Code ORDER BY LeaveID DESC";

                    using (var command = new SqliteCommand(query, connection))
                    {
                        if (!string.IsNullOrEmpty(_employeeCode))
                        {
                            command.Parameters.AddWithValue("@Code", _employeeCode);
                        }

                        List<LeaveHistoryItem> leaveList = new List<LeaveHistoryItem>();

                        using (var reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                string actualReturn = reader["ActualReturnDate"].ToString();
                                string expDateStr = reader["ExpectedReturnDate"].ToString();

                                int lateDays = 0;
                                if (reader["LateDays"] != DBNull.Value && int.TryParse(reader["LateDays"].ToString(), out int parsedLate))
                                {
                                    lateDays = parsedLate;
                                }

                                if (string.IsNullOrWhiteSpace(actualReturn) && DateTime.TryParse(expDateStr, out DateTime expectedReturnDT))
                                {
                                    if (expectedReturnDT < DateTime.Today)
                                    {
                                        lateDays = (DateTime.Today - expectedReturnDT).Days;
                                    }
                                }

                                string statusColor = "#27AE60"; // Green
                                if (lateDays > 7)
                                    statusColor = "#FF5252"; // Red 
                                else if (lateDays > 0)
                                    statusColor = "#F1C40F"; // Yellow 

                                leaveList.Add(new LeaveHistoryItem
                                {
                                    LeaveID = reader["LeaveID"].ToString(),
                                    DepartureDate = reader["DepartureDate"].ToString(),
                                    ExpectedReturnDate = expDateStr,
                                    ActualReturnDate = actualReturn,
                                    LateDaysText = lateDays.ToString(),
                                    Reason = reader["Reason"].ToString(),
                                    StatusColor = statusColor
                                });
                            }
                        }

                        dgLeaves.ItemsSource = leaveList;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading leave history: {ex.Message}");
            }
        }

        private void btnEditLeave_Click(object sender, RoutedEventArgs e)
        {
            var button = sender as Button;
            var item = button.DataContext as LeaveHistoryItem;

            if (item != null)
            {
                _editingLeaveId = item.LeaveID;

                dpDeparture.SelectedDate = string.IsNullOrEmpty(item.DepartureDate) ? (DateTime?)null : DateTime.Parse(item.DepartureDate);
                dpExpected.SelectedDate = string.IsNullOrEmpty(item.ExpectedReturnDate) ? (DateTime?)null : DateTime.Parse(item.ExpectedReturnDate);
                dpActual.SelectedDate = string.IsNullOrEmpty(item.ActualReturnDate) ? (DateTime?)null : DateTime.Parse(item.ActualReturnDate);
                txtReason.Text = item.Reason;

                btnSaveLeave.Content = "Update Leave Record";
                btnSaveLeave.Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#27AE60"));
                btnCancelEdit.Visibility = Visibility.Visible;
            }
        }

        private void btnCancelEdit_Click(object sender, RoutedEventArgs e)
        {
            ResetForm();
        }

        private void ResetForm()
        {
            _editingLeaveId = null;
            dpDeparture.SelectedDate = null;
            dpExpected.SelectedDate = null;
            dpActual.SelectedDate = null;
            txtReason.Text = "";

            btnSaveLeave.Content = "Save Leave Record";
            btnSaveLeave.Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#0078D7"));
            btnCancelEdit.Visibility = Visibility.Collapsed;
        }

        private void btnSaveLeave_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(_employeeCode))
            {
                MessageBox.Show("Cannot add leaves from the Global View. Please open a specific employee's details window to add or update leave records.", "Action Not Allowed", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (dpDeparture.SelectedDate == null || dpExpected.SelectedDate == null)
            {
                MessageBox.Show("Please select at least the Departure Date and Expected Return Date.", "Validation Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            try
            {
                int lateDays = 0;
                string actualReturnStr = "";

                if (dpActual.SelectedDate.HasValue)
                {
                    actualReturnStr = dpActual.SelectedDate.Value.ToString("yyyy-MM-dd");
                    TimeSpan difference = dpActual.SelectedDate.Value - dpExpected.SelectedDate.Value;
                    if (difference.Days > 0)
                    {
                        lateDays = difference.Days;
                    }
                }

                using (var connection = new SqliteConnection(DatabaseHelper.GetConnectionString()))
                {
                    connection.Open();

                    string ensureProfileQuery = "INSERT OR IGNORE INTO EmployeeProfiles (EmployeeCode) VALUES (@Code)";
                    using (var ensureCmd = new SqliteCommand(ensureProfileQuery, connection))
                    {
                        ensureCmd.Parameters.AddWithValue("@Code", _employeeCode);
                        ensureCmd.ExecuteNonQuery();
                    }

                    string query;

                    if (string.IsNullOrEmpty(_editingLeaveId))
                    {
                        query = @"
                            INSERT INTO LongLeaveTracking (EmployeeCode, DepartureDate, ExpectedReturnDate, ActualReturnDate, LateDays, Reason) 
                            VALUES (@Code, @Departure, @Expected, @Actual, @LateDays, @Reason)";
                    }
                    else
                    {
                        query = @"
                            UPDATE LongLeaveTracking 
                            SET DepartureDate = @Departure, ExpectedReturnDate = @Expected, ActualReturnDate = @Actual, LateDays = @LateDays, Reason = @Reason
                            WHERE LeaveID = @LeaveID";
                    }

                    using (var command = new SqliteCommand(query, connection))
                    {
                        command.Parameters.AddWithValue("@Code", _employeeCode);
                        command.Parameters.AddWithValue("@Departure", dpDeparture.SelectedDate.Value.ToString("yyyy-MM-dd"));
                        command.Parameters.AddWithValue("@Expected", dpExpected.SelectedDate.Value.ToString("yyyy-MM-dd"));
                        command.Parameters.AddWithValue("@Actual", actualReturnStr);
                        command.Parameters.AddWithValue("@LateDays", lateDays);
                        command.Parameters.AddWithValue("@Reason", txtReason.Text ?? "");

                        if (!string.IsNullOrEmpty(_editingLeaveId))
                        {
                            command.Parameters.AddWithValue("@LeaveID", _editingLeaveId);
                        }

                        command.ExecuteNonQuery();
                    }
                }

                string successMessage = string.IsNullOrEmpty(_editingLeaveId) ? "Leave record saved successfully!" : "Leave record updated successfully!";
                MessageBox.Show(successMessage, "Success", MessageBoxButton.OK, MessageBoxImage.Information);

                ResetForm();
                LoadLeaveHistory();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error saving leave record: {ex.Message}", "Database Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void btnDeleteLeave_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(_employeeCode))
            {
                MessageBox.Show("Cannot delete leaves from the Global View. Please open the specific employee's details window.", "Action Not Allowed", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var button = sender as Button;
            var item = button.DataContext as LeaveHistoryItem;

            if (item != null)
            {
                string leaveId = item.LeaveID;

                MessageBoxResult result = MessageBox.Show("Are you sure you want to delete this leave record?", "Confirm Delete", MessageBoxButton.YesNo, MessageBoxImage.Warning);

                if (result == MessageBoxResult.Yes)
                {
                    try
                    {
                        using (var connection = new SqliteConnection(DatabaseHelper.GetConnectionString()))
                        {
                            connection.Open();
                            string query = "DELETE FROM LongLeaveTracking WHERE LeaveID = @LeaveID";

                            using (var command = new SqliteCommand(query, connection))
                            {
                                command.Parameters.AddWithValue("@LeaveID", leaveId);
                                command.ExecuteNonQuery();
                            }
                        }

                        if (_editingLeaveId == leaveId) ResetForm();
                        LoadLeaveHistory();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Error deleting record: {ex.Message}", "Delete Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
            }
        }

        private void btnImportLegacy_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Title = "Select Raw Excel CSV Data",
                Filter = "CSV Files (*.csv)|*.csv"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                try
                {
                    string[] lines = File.ReadAllLines(openFileDialog.FileName);
                    int importCount = 0;
                    int failCount = 0;

                    int cardIdx = 0, depIdx = 1, expIdx = 2, actIdx = 3, noteIdx = 4;
                    int startRow = 1;

                    // 🔥 SMART SCANNER: Looks through the first 10 rows to find exact Excel Headers
                    // Ensures strict check for "ACTUAL" before "RETURN" to prevent mapping both to the same column
                    for (int i = 0; i < Math.Min(10, lines.Length); i++)
                    {
                        string upperLine = lines[i].ToUpper();
                        if (upperLine.Contains("CARD") || upperLine.Contains("DEPARTURE"))
                        {
                            var headers = lines[i].Split(',');
                            for (int c = 0; c < headers.Length; c++)
                            {
                                string h = headers[c].ToUpper().Trim();
                                if (h.Contains("CARD")) cardIdx = c;
                                else if (h.Contains("DEPARTURE")) depIdx = c;
                                else if (h.Contains("ACTUAL") || h == "DAYS" || h == "DAY") actIdx = c;
                                else if (h.Contains("EXPECTED") || h.Contains("RETURN")) expIdx = c;
                                else if (h.Contains("REMARK") || h.Contains("NOTE")) noteIdx = c;
                            }
                            startRow = i + 1;
                            break;
                        }
                    }

                    using (var connection = new SqliteConnection(DatabaseHelper.GetConnectionString()))
                    {
                        connection.Open();

                        for (int i = startRow; i < lines.Length; i++)
                        {
                            string line = lines[i];
                            if (string.IsNullOrWhiteSpace(line)) continue;

                            string[] cols = line.Split(',');

                            // Failsafe: Ensure the row actually has enough columns
                            if (cols.Length <= Math.Max(cardIdx, Math.Max(depIdx, expIdx))) continue;

                            string rawCard = cols[cardIdx].Trim();

                            // Skip if the Card column is empty or just random Excel text
                            if (string.IsNullOrWhiteSpace(rawCard) || !int.TryParse(rawCard, out _)) continue;

                            // Automatically format employee code to 4 digits
                            string empCode = rawCard.PadLeft(4, '0');

                            DateTime? ParseDate(string d)
                            {
                                if (string.IsNullOrWhiteSpace(d)) return null;
                                string cleanDate = d.Trim();
                                string[] formats = {
                                    "dd-MM-yyyy", "d-M-yyyy", "dd/MM/yyyy", "M/d/yyyy",
                                    "dd-MM-yy", "d-M-yy", "yyyy-MM-dd"
                                };
                                if (DateTime.TryParseExact(cleanDate, formats, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime dt))
                                    return dt;
                                return null;
                            }

                            DateTime? depDate = ParseDate(cols[depIdx]);
                            DateTime? expReturn = ParseDate(cols[expIdx]);

                            // Safely check for Actual Return Date in the "DAYS" column
                            DateTime? actReturn = null;
                            if (cols.Length > actIdx && actIdx != expIdx)
                            {
                                actReturn = ParseDate(cols[actIdx]);
                            }

                            string notes = cols.Length > noteIdx ? cols[noteIdx].Trim() : "";

                            // We must have at least a Departure and Expected Return to save it
                            if (depDate.HasValue && expReturn.HasValue)
                            {
                                int lateDays = 0;
                                if (actReturn.HasValue)
                                {
                                    if ((actReturn.Value - expReturn.Value).Days > 0)
                                        lateDays = (actReturn.Value - expReturn.Value).Days;
                                }
                                else if (expReturn.Value < DateTime.Today)
                                {
                                    lateDays = (DateTime.Today - expReturn.Value).Days;
                                }

                                string sql = @"INSERT INTO LongLeaveTracking 
                                             (EmployeeCode, DepartureDate, ExpectedReturnDate, ActualReturnDate, LateDays, Reason) 
                                             VALUES (@EmpCode, @Dep, @Exp, @Act, @LateDays, @Reason)";

                                using (var cmd = new SqliteCommand(sql, connection))
                                {
                                    cmd.Parameters.AddWithValue("@EmpCode", empCode);
                                    cmd.Parameters.AddWithValue("@Dep", depDate.Value.ToString("yyyy-MM-dd"));
                                    cmd.Parameters.AddWithValue("@Exp", expReturn.Value.ToString("yyyy-MM-dd"));
                                    cmd.Parameters.AddWithValue("@Act", actReturn.HasValue ? actReturn.Value.ToString("yyyy-MM-dd") : "");
                                    cmd.Parameters.AddWithValue("@LateDays", lateDays);
                                    cmd.Parameters.AddWithValue("@Reason", notes);

                                    // Prevent relational database crashes by ensuring the employee profile exists
                                    using (var profileCmd = new SqliteCommand("INSERT OR IGNORE INTO EmployeeProfiles (EmployeeCode) VALUES (@EmpCode)", connection))
                                    {
                                        profileCmd.Parameters.AddWithValue("@EmpCode", empCode);
                                        profileCmd.ExecuteNonQuery();
                                    }

                                    cmd.ExecuteNonQuery();
                                    importCount++;
                                }
                            }
                            else
                            {
                                failCount++;
                            }
                        }
                    }

                    if (failCount > 0)
                    {
                        MessageBox.Show($"Successfully imported {importCount} records.\n\nSkipped {failCount} messy rows (Departure or Return dates were blank or corrupted).", "Import Completed with Skips", MessageBoxButton.OK, MessageBoxImage.Warning);
                    }
                    else if (importCount > 0)
                    {
                        MessageBox.Show($"Successfully imported {importCount} historical leave records!", "Import Complete", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                    else
                    {
                        MessageBox.Show("No records could be imported. Please ensure you saved the file as a Comma Delimited CSV.", "Import Failed", MessageBoxButton.OK, MessageBoxImage.Error);
                    }

                    LoadLeaveHistory();
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error reading CSV file. Ensure it is closed in Excel before importing.\n\nDetails: {ex.Message}", "File Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }
    }
}