using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using Microsoft.Win32;
using Microsoft.Data.Sqlite;
using ArisPayrollCompanion.Helpers;
using ArisPayrollCompanion.Views;

namespace ArisPayrollCompanion
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            txtFolder.Text = App.DbfFolderPath;
            txtStatus.Text = "Ready";
        }

        private void btnBrowse_Click(object sender, RoutedEventArgs e)
        {
            using (var dialog = new System.Windows.Forms.FolderBrowserDialog())
            {
                dialog.Description = "Select the DBFS folder containing your FoxPro data";
                dialog.SelectedPath = App.DbfFolderPath;

                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    txtFolder.Text = dialog.SelectedPath;
                    App.DbfFolderPath = dialog.SelectedPath;
                    txtStatus.Text = "Path updated and saved successfully.";
                }
            }
        }

        private void btnEmployees_Click(object sender, RoutedEventArgs e)
        {
            var employeeListWindow = new EmployeeListWindow();
            employeeListWindow.Owner = this;
            employeeListWindow.ShowDialog();
        }

        private void btnSync_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (!Directory.Exists(App.DbfFolderPath))
                {
                    MessageBox.Show("FoxPro folder not found: " + App.DbfFolderPath, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                txtStatus.Text = "Syncing employees... Please wait.";

                var readerService = new ArisPayrollCompanion.Services.DbfReaderService();
                string dummyPath = Path.Combine(App.DbfFolderPath, "EMPMAS.DBF");
                var employees = readerService.ReadEmployees(dummyPath);

                if (employees == null || employees.Count == 0)
                {
                    MessageBox.Show("No employees could be extracted from the DBF files.", "Sync Result", MessageBoxButton.OK, MessageBoxImage.Warning);
                    txtStatus.Text = "Ready";
                    return;
                }

                int newCount = 0;
                int updateCount = 0;

                using (var connection = new SqliteConnection(DatabaseHelper.GetConnectionString()))
                {
                    connection.Open();

                    string[] cols = { "EmployeeName TEXT", "CardNumber TEXT", "Department TEXT", "Designation TEXT", "Category TEXT", "Status TEXT" };
                    foreach (var col in cols)
                    {
                        try
                        {
                            using (var cmd = new SqliteCommand($"ALTER TABLE EmployeeProfiles ADD COLUMN {col}", connection))
                            {
                                cmd.ExecuteNonQuery();
                            }
                        }
                        catch { /* Already exists */ }
                    }

                    foreach (var emp in employees)
                    {
                        using (var checkCmd = new SqliteCommand("SELECT COUNT(1) FROM EmployeeProfiles WHERE EmployeeCode = @Code", connection))
                        {
                            checkCmd.Parameters.AddWithValue("@Code", emp.EmployeeCode);
                            long exists = (long)checkCmd.ExecuteScalar();

                            if (exists > 0)
                            {
                                string updateSql = @"UPDATE EmployeeProfiles 
                                                     SET EmployeeName = @Name, CardNumber = @Card, Department = @Dept, 
                                                         Designation = @Desig, Category = @Cat, Status = @Status 
                                                     WHERE EmployeeCode = @Code";
                                using (var updateCmd = new SqliteCommand(updateSql, connection))
                                {
                                    updateCmd.Parameters.AddWithValue("@Code", emp.EmployeeCode);
                                    updateCmd.Parameters.AddWithValue("@Name", emp.EmployeeName ?? "");
                                    updateCmd.Parameters.AddWithValue("@Card", emp.CardNumber ?? "");
                                    updateCmd.Parameters.AddWithValue("@Dept", emp.Department ?? "");
                                    updateCmd.Parameters.AddWithValue("@Desig", emp.Designation ?? "");
                                    updateCmd.Parameters.AddWithValue("@Cat", emp.Category ?? "");
                                    updateCmd.Parameters.AddWithValue("@Status", emp.Status ?? "Y");
                                    updateCmd.ExecuteNonQuery();
                                }
                                updateCount++;
                            }
                            else
                            {
                                string insertSql = @"INSERT INTO EmployeeProfiles 
                                                     (EmployeeCode, EmployeeName, CardNumber, Department, Designation, Category, Status) 
                                                     VALUES (@Code, @Name, @Card, @Dept, @Desig, @Cat, @Status)";
                                using (var insertCmd = new SqliteCommand(insertSql, connection))
                                {
                                    insertCmd.Parameters.AddWithValue("@Code", emp.EmployeeCode);
                                    insertCmd.Parameters.AddWithValue("@Name", emp.EmployeeName ?? "");
                                    insertCmd.Parameters.AddWithValue("@Card", emp.CardNumber ?? "");
                                    insertCmd.Parameters.AddWithValue("@Dept", emp.Department ?? "");
                                    insertCmd.Parameters.AddWithValue("@Desig", emp.Designation ?? "");
                                    insertCmd.Parameters.AddWithValue("@Cat", emp.Category ?? "");
                                    insertCmd.Parameters.AddWithValue("@Status", emp.Status ?? "Y");
                                    insertCmd.ExecuteNonQuery();
                                }
                                newCount++;
                            }
                        }
                    }
                }

                txtStatus.Text = "Ready";
                MessageBox.Show($"Sync Complete!\n\nAdded {newCount} new employees.\nUpdated {updateCount} existing employees.\n\nTotal in database: {employees.Count}",
                                "Sync Successful", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Critical error syncing employees: {ex.Message}", "Sync Error", MessageBoxButton.OK, MessageBoxImage.Error);
                txtStatus.Text = "Ready";
            }
        }

        private DataTable ReadDbfHeaderOnly(string filePath)
        {
            DataTable dt = new DataTable();
            if (!File.Exists(filePath)) return dt;

            try
            {
                using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (BinaryReader br = new BinaryReader(fs))
                {
                    br.ReadBytes(4);
                    br.ReadInt32();
                    int headerLen = br.ReadInt16();
                    br.ReadInt16();
                    br.ReadBytes(20);

                    while (br.BaseStream.Position < headerLen - 1)
                    {
                        byte firstByte = br.ReadByte();
                        if (firstByte == 0x0D) break;
                        br.BaseStream.Position--;

                        byte[] nameBytes = br.ReadBytes(11);
                        string name = Encoding.ASCII.GetString(nameBytes).Replace("\0", "").Trim();

                        br.ReadByte(); br.ReadBytes(4);
                        br.ReadByte();
                        br.ReadBytes(15);

                        if (!string.IsNullOrWhiteSpace(name))
                        {
                            dt.Columns.Add(name.ToUpper());
                        }
                    }
                }
            }
            catch { }
            return dt;
        }

        private void btnTestConnection_Click(object sender, RoutedEventArgs e)
        {
            string filePath = Path.Combine(App.DbfFolderPath, "monsum.DBF");

            if (!File.Exists(filePath))
            {
                MessageBox.Show("File not found: " + filePath);
                return;
            }

            try
            {
                DataTable dt = ReadDbfHeaderOnly(filePath);
                StringBuilder sb = new StringBuilder();
                sb.AppendLine("All Columns Found in monsum.DBF:");
                sb.AppendLine("--------------------------------");

                foreach (DataColumn col in dt.Columns)
                {
                    sb.AppendLine(col.ColumnName);
                }

                MessageBox.Show(sb.ToString(), "Full Column List");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error reading file: " + ex.Message);
            }
        }

        private void btnActiveLeaves_Click(object sender, RoutedEventArgs e)
        {
            new LeaveTrackerWindow("", "") { Owner = this }.ShowDialog();
        }

        private void btnActiveLoansReport_Click(object sender, RoutedEventArgs e)
        {
            new ActiveLoansReportWindow { Owner = this }.ShowDialog();
        }

        private void btnCompliance_Click(object sender, RoutedEventArgs e)
        {
            new ComplianceDashboardWindow { Owner = this }.ShowDialog();
        }

        private void btnInactiveLoansReport_Click(object sender, RoutedEventArgs e)
        {
            new InactiveLoansReportWindow { Owner = this }.ShowDialog();
        }

        private void BtnImportAdvanceLimits_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Title = "Select Advance Limits CSV",
                Filter = "CSV Files (*.csv)|*.csv"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                try
                {
                    string[] lines = File.ReadAllLines(openFileDialog.FileName);
                    int importCount = 0;
                    int failCount = 0;

                    int cardIdx = 1, advIdx = 3;
                    int startRow = 1;

                    for (int i = 0; i < Math.Min(10, lines.Length); i++)
                    {
                        string upperLine = lines[i].ToUpper();
                        if (upperLine.Contains("CARD") && upperLine.Contains("ADV"))
                        {
                            var headers = lines[i].Split(',');
                            for (int c = 0; c < headers.Length; c++)
                            {
                                string h = headers[c].ToUpper().Trim();
                                if (h.Contains("CARD")) cardIdx = c;
                                else if (h.Contains("ADV")) advIdx = c;
                            }
                            startRow = i + 1;
                            break;
                        }
                    }

                    string baseDir = AppDomain.CurrentDomain.BaseDirectory;

                    for (int i = startRow; i < lines.Length; i++)
                    {
                        string line = lines[i];
                        if (string.IsNullOrWhiteSpace(line)) continue;

                        string[] cols = line.Split(',');

                        if (cols.Length <= Math.Max(cardIdx, advIdx)) continue;

                        string rawCard = cols[cardIdx].Trim();
                        string rawAdv = cols[advIdx].Trim();

                        if (string.IsNullOrWhiteSpace(rawCard) || !int.TryParse(rawCard, out _)) continue;

                        string empCode = rawCard.PadLeft(4, '0');

                        if (decimal.TryParse(rawAdv, out decimal limitValue))
                        {
                            string limitFilePath = Path.Combine(baseDir, $"advance_limit_{empCode}.txt");
                            File.WriteAllText(limitFilePath, limitValue.ToString("0"));

                            importCount++;
                        }
                        else
                        {
                            failCount++;
                        }
                    }

                    if (failCount > 0)
                    {
                        MessageBox.Show($"Successfully updated {importCount} employee limits!\n\nSkipped {failCount} rows (ADV limit was empty or contained text).", "Import Completed with Skips", MessageBoxButton.OK, MessageBoxImage.Warning);
                    }
                    else if (importCount > 0)
                    {
                        MessageBox.Show($"Successfully updated {importCount} employee advance limits!", "Import Complete", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                    else
                    {
                        MessageBox.Show("No records could be imported. Please ensure the file is saved as a CSV and contains 'Card No' and 'ADV' headers.", "Import Failed", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error reading CSV file. Ensure it is closed in Excel before importing.\n\nDetails: {ex.Message}", "File Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void BtnDailyReports_Click(object sender, RoutedEventArgs e)
        {
            new DailyReportsWindow { Owner = this }.ShowDialog();
        }

        // 🔥 NEW: Employee History Report Click Event
        private void BtnHistoryReport_Click(object sender, RoutedEventArgs e)
        {
            new EmployeeHistoryReportWindow { Owner = this }.ShowDialog();
        }
    }
}