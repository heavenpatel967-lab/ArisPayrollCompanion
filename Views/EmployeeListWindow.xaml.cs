using ArisPayrollCompanion.Models;
using ArisPayrollCompanion.Helpers;
using Microsoft.Data.Sqlite;
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
using System.Windows.Input;

namespace ArisPayrollCompanion.Views
{
    public partial class EmployeeListWindow : Window
    {
        private bool _isLoaded = false;
        private Dictionary<string, (string DeptCode, string CatCode)> _monsumData;
        private string _dbfFolderPath = "";

        public EmployeeListWindow()
        {
            InitializeComponent();

            try
            {
                string dbPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Payroll.db");

                if (!File.Exists(dbPath))
                {
                    MessageBox.Show($"Database file missing!\nExpected at: {dbPath}", "Database Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    this.Loaded += (s, e) => this.Close();
                    return;
                }

                string savedPath = App.DbfFolderPath;

                if (!string.IsNullOrEmpty(savedPath) && Directory.Exists(savedPath))
                {
                    _dbfFolderPath = savedPath;
                }
                else
                {
                    MessageBox.Show("Payroll data path is missing. Please go to the Main Window and select the correct DBF folder.", "Path Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                    this.Loaded += (s, e) => this.Close();
                    return;
                }

                LoadMonsumData();
                LoadFilters();
                _isLoaded = true;
                LoadEmployees();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"CRITICAL ERROR: {ex.Message}", "Initialization Failed", MessageBoxButton.OK, MessageBoxImage.Error);
                this.Loaded += (s, e) => this.Close();
            }
        }

        private void LoadMonsumData()
        {
            _monsumData = new Dictionary<string, (string, string)>();
            string filePath = Path.Combine(_dbfFolderPath, "monsum.DBF");

            if (!File.Exists(filePath))
            {
                filePath = Path.Combine(_dbfFolderPath, "monsum.dbf");
                if (!File.Exists(filePath)) return;
            }

            try
            {
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
                        if (!string.IsNullOrEmpty(name)) fields.Add(new Tuple<string, int>(name, len));
                    }
                    reader.BaseStream.Position = headerSize;

                    int empCdIdx = fields.FindIndex(f => f.Item1 == "EMPCD");
                    int deptIdx = fields.FindIndex(f => f.Item1 == "DEPT");
                    int catIdx = fields.FindIndex(f => f.Item1 == "CAT");

                    if (empCdIdx == -1) return;

                    for (int i = 0; i < numRecords; i++)
                    {
                        byte[] rowData = reader.ReadBytes(recordSize);
                        if (rowData.Length < recordSize) break;
                        string empCode = ExtractField(rowData, fields, empCdIdx);
                        string dept = deptIdx != -1 ? ExtractField(rowData, fields, deptIdx) : "";
                        string cat = catIdx != -1 ? ExtractField(rowData, fields, catIdx) : "";

                        if (!string.IsNullOrWhiteSpace(empCode))
                        {
                            if (!_monsumData.ContainsKey(empCode)) _monsumData[empCode] = (dept, cat);
                            else
                            {
                                var existing = _monsumData[empCode];
                                _monsumData[empCode] = (!string.IsNullOrWhiteSpace(dept) ? dept : existing.DeptCode,
                                                       !string.IsNullOrWhiteSpace(cat) ? cat : existing.CatCode);
                            }
                        }
                    }
                }
            }
            catch { }
        }

        private string ExtractField(byte[] rowData, List<Tuple<string, int>> fields, int index)
        {
            int offset = 1;
            for (int i = 0; i < index; i++) offset += fields[i].Item2;
            if (offset + fields[index].Item2 <= rowData.Length)
                return Encoding.Latin1.GetString(rowData, offset, fields[index].Item2).Trim();
            return "";
        }

        // 🔥 FIX: Bypasses the restricted repository and pulls distinct values straight from the DB
        private void LoadFilters()
        {
            List<string> rawDepts = new List<string>();
            List<string> rawCats = new List<string>();
            List<string> rawDesigs = new List<string>();

            using (var conn = new SqliteConnection(DatabaseHelper.GetConnectionString()))
            {
                conn.Open();
                using (var cmd = new SqliteCommand("SELECT DISTINCT Department FROM EmployeeProfiles WHERE Department IS NOT NULL AND Department != ''", conn))
                using (var reader = cmd.ExecuteReader()) { while (reader.Read()) rawDepts.Add(reader.GetString(0)); }

                using (var cmd = new SqliteCommand("SELECT DISTINCT Category FROM EmployeeProfiles WHERE Category IS NOT NULL AND Category != ''", conn))
                using (var reader = cmd.ExecuteReader()) { while (reader.Read()) rawCats.Add(reader.GetString(0)); }

                using (var cmd = new SqliteCommand("SELECT DISTINCT Designation FROM EmployeeProfiles WHERE Designation IS NOT NULL AND Designation != ''", conn))
                using (var reader = cmd.ExecuteReader()) { while (reader.Read()) rawDesigs.Add(reader.GetString(0)); }
            }

            var translatedDepts = rawDepts.Where(d => d != "All Departments").Select(d => GetDepartmentName(d)).Distinct().OrderBy(d => d).ToList();
            translatedDepts.Insert(0, "All Departments");
            cmbDepartment.ItemsSource = translatedDepts;
            cmbDepartment.SelectedIndex = 0;

            var translatedCats = rawCats.Where(c => c != "All Categories").Select(c => GetCategoryName(c)).Distinct().OrderBy(c => c).ToList();
            translatedCats.Insert(0, "All Categories");
            cmbCategory.ItemsSource = translatedCats;
            cmbCategory.SelectedIndex = 0;

            var translatedDesigs = rawDesigs.OrderBy(d => d).ToList();
            translatedDesigs.Insert(0, "All Designations");
            cmbDesignation.ItemsSource = translatedDesigs;
            cmbDesignation.SelectedIndex = 0;

            cmbStatus.ItemsSource = new List<string> { "All Statuses", "Y", "N" };
            cmbStatus.SelectedIndex = 0;
        }

        private void txtSearch_TextChanged(object sender, TextChangedEventArgs e) => LoadEmployees();
        private void Filter_SelectionChanged(object sender, SelectionChangedEventArgs e) => LoadEmployees();

        // 🔥 FIX: Reads directly from SQLite, completely ignoring any hidden repository restrictions
        private void LoadEmployees()
        {
            if (!_isLoaded) return;

            var allEmployees = new List<Employee>();

            using (var conn = new SqliteConnection(DatabaseHelper.GetConnectionString()))
            {
                conn.Open();
                string query = "SELECT EmployeeCode, EmployeeName, CardNumber, Department, Designation, Category, Status FROM EmployeeProfiles";

                using (var cmd = new SqliteCommand(query, conn))
                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        allEmployees.Add(new Employee
                        {
                            EmployeeCode = reader.IsDBNull(0) ? "" : reader.GetString(0),
                            EmployeeName = reader.IsDBNull(1) ? "" : reader.GetString(1),
                            CardNumber = reader.IsDBNull(2) ? "" : reader.GetString(2),
                            Department = reader.IsDBNull(3) ? "" : reader.GetString(3),
                            Designation = reader.IsDBNull(4) ? "" : reader.GetString(4),
                            Category = reader.IsDBNull(5) ? "" : reader.GetString(5),
                            Status = reader.IsDBNull(6) ? "Y" : reader.GetString(6)
                        });
                    }
                }
            }

            foreach (var emp in allEmployees)
            {
                if (_monsumData != null && _monsumData.TryGetValue(emp.EmployeeCode, out var trueCodes))
                {
                    emp.Department = !string.IsNullOrWhiteSpace(trueCodes.DeptCode) ? GetDepartmentName(trueCodes.DeptCode) : emp.Department;
                    emp.Category = !string.IsNullOrWhiteSpace(trueCodes.CatCode) ? GetCategoryName(trueCodes.CatCode) : emp.Category;
                }

                string rawStatus = emp.Status?.Trim().ToUpper() ?? "";
                emp.Status = (rawStatus == "N" || rawStatus == "FALSE" || rawStatus == "0") ? "N" : "Y";
            }

            string searchText = txtSearch.Text.ToLower().Trim();
            string selectedDept = cmbDepartment.SelectedItem as string ?? "All Departments";
            string selectedCat = cmbCategory.SelectedItem as string ?? "All Categories";
            string selectedDesig = cmbDesignation.SelectedItem as string ?? "All Designations";
            string selectedStatus = cmbStatus.SelectedItem as string ?? "All Statuses";

            var filtered = allEmployees.Where(e =>
                (string.IsNullOrEmpty(searchText) || e.EmployeeName.ToLower().Contains(searchText) || e.EmployeeCode.Contains(searchText) || e.CardNumber.Contains(searchText)) &&
                (selectedDept == "All Departments" || e.Department == selectedDept) &&
                (selectedCat == "All Categories" || e.Category == selectedCat) &&
                (selectedDesig == "All Designations" || e.Designation == selectedDesig) &&
                (selectedStatus == "All Statuses" || e.Status == selectedStatus)
            ).ToList();

            dgEmployees.ItemsSource = filtered.OrderBy(e => int.TryParse(e.EmployeeCode, out int n) ? n : int.MaxValue).ToList();
            txtTotal.Text = $"Total Employees : {filtered.Count}";
        }

        private void dgEmployees_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (dgEmployees.SelectedItem is Employee emp)
            {
                new EmployeeDetailsWindow(emp) { Owner = this }.ShowDialog();
                LoadEmployees();
            }
        }

        private void btnExportExcel_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var currentList = dgEmployees.ItemsSource as List<Employee>;
                if (currentList == null || currentList.Count == 0) return;
                SaveFileDialog sfd = new SaveFileDialog { Filter = "Excel Workbook|*.xlsx", FileName = $"Export_{DateTime.Now:yyyyMMdd}.xlsx" };
                if (sfd.ShowDialog() == true)
                {
                    using (var wb = new XLWorkbook()) { var ws = wb.Worksheets.Add("Employees"); ws.Cell(1, 1).InsertTable(currentList); wb.SaveAs(sfd.FileName); }
                    MessageBox.Show("Export complete!");
                }
            }
            catch (Exception ex) { MessageBox.Show($"Error: {ex.Message}"); }
        }

        private string GetDepartmentName(string code) => code?.Trim().TrimStart('0') switch { "1" => "OFFICE", "2" => "CASTING", "3" => "MACHINE", "4" => "BUFF & AMRI", "5" => "CHROM", "6" => "FITTING / OTHER", "7" => "OTHER NEW", _ => code };
        private string GetCategoryName(string code) => code?.Trim().TrimStart('0') switch { "1" => "STAFF", "2" => "CASTING_10", "3" => "WORKER", "4" => "MONTHLY_8", "5" => "GENERAL / PIC RATE", "6" => "GENERAL ODD", "7" => "MONTHLY_10", "8" => "TRANSPORT", "10" => "PLUMBER", "11" => "PIC RATE GENERAL", "12" => "WORKER_9", _ => code };
    }
}