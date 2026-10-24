using System;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;
using Microsoft.Data.Sqlite;
using ArisPayrollCompanion.Helpers;

namespace ArisPayrollCompanion.Views
{
    public partial class EmployeeDocumentsWindow : Window
    {
        private readonly string _employeeCode;
        private string _selectedFilePath = "";

        public EmployeeDocumentsWindow(string employeeCode, string employeeName)
        {
            InitializeComponent();
            _employeeCode = employeeCode;
            txtHeader.Text = $"Documents Manager: {employeeName} ({employeeCode})";
            LoadDocuments();
        }

        private void LoadDocuments()
        {
            try
            {
                using (var connection = new SqliteConnection(DatabaseHelper.GetConnectionString()))
                {
                    connection.Open();
                    string query = "SELECT DocID, DocumentCategory, DocumentName, FilePath, UploadDate FROM EmployeeDocuments WHERE EmployeeCode = @Code ORDER BY DocID DESC";

                    using (var command = new SqliteCommand(query, connection))
                    {
                        command.Parameters.AddWithValue("@Code", _employeeCode);

                        DataTable dt = new DataTable();
                        using (var reader = command.ExecuteReader())
                        {
                            dt.Load(reader);
                        }

                        dgDocuments.ItemsSource = dt.DefaultView;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading documents: {ex.Message}");
            }
        }

        private void btnBrowse_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Title = "Select Document",
                Filter = "All Supported Files|*.pdf;*.jpg;*.jpeg;*.png|PDF Files|*.pdf|Images|*.jpg;*.jpeg;*.png"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                _selectedFilePath = openFileDialog.FileName;
                txtSelectedFile.Text = $"Selected: {Path.GetFileName(_selectedFilePath)}";
            }
        }

        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(_selectedFilePath) || string.IsNullOrWhiteSpace(txtDocName.Text))
            {
                MessageBox.Show("Please select a file and enter a document name.", "Validation Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            try
            {
                // 1. Define the safe folder inside your application directory
                string baseAppFolder = AppDomain.CurrentDomain.BaseDirectory;
                string vaultFolder = Path.Combine(baseAppFolder, "CompanionVault", "Documents", _employeeCode);

                // 2. Create the folder if it doesn't exist yet
                if (!Directory.Exists(vaultFolder))
                {
                    Directory.CreateDirectory(vaultFolder);
                }

                // 3. Create a unique file name (Timestamp + Original Name)
                string originalFileName = Path.GetFileName(_selectedFilePath);
                string uniqueFileName = DateTime.Now.ToString("yyyyMMdd_HHmmss_") + originalFileName;
                string secureFilePath = Path.Combine(vaultFolder, uniqueFileName);

                // 4. COPY the file from the user's PC into the App's secure vault
                File.Copy(_selectedFilePath, secureFilePath, true);

                using (var connection = new SqliteConnection(DatabaseHelper.GetConnectionString()))
                {
                    connection.Open();

                    // Ensure the employee exists in the Profiles table first!
                    string ensureProfileQuery = "INSERT OR IGNORE INTO EmployeeProfiles (EmployeeCode) VALUES (@Code)";
                    using (var ensureCmd = new SqliteCommand(ensureProfileQuery, connection))
                    {
                        ensureCmd.Parameters.AddWithValue("@Code", _employeeCode);
                        ensureCmd.ExecuteNonQuery();
                    }

                    // Proceed with saving the document using the secure path
                    string query = @"
                        INSERT INTO EmployeeDocuments (EmployeeCode, DocumentCategory, DocumentName, FilePath, UploadDate) 
                        VALUES (@Code, @Cat, @Name, @Path, @Date)";

                    using (var command = new SqliteCommand(query, connection))
                    {
                        command.Parameters.AddWithValue("@Code", _employeeCode);
                        command.Parameters.AddWithValue("@Cat", (cmbCategory.SelectedItem as ComboBoxItem)?.Content.ToString());
                        command.Parameters.AddWithValue("@Name", txtDocName.Text);
                        command.Parameters.AddWithValue("@Path", secureFilePath);
                        command.Parameters.AddWithValue("@Date", DateTime.Now.ToString("yyyy-MM-dd"));

                        command.ExecuteNonQuery();
                    }
                }

                MessageBox.Show("Document copied and saved securely to the App Vault!", "Success", MessageBoxButton.OK, MessageBoxImage.Information);

                // Reset fields and reload table
                _selectedFilePath = "";
                txtSelectedFile.Text = "";
                txtDocName.Text = "";
                cmbCategory.SelectedIndex = 0;
                LoadDocuments();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error saving document: {ex.Message}", "System Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void btnOpenDocument_Click(object sender, RoutedEventArgs e)
        {
            var button = sender as Button;
            var dataRowView = button.DataContext as DataRowView;

            if (dataRowView != null)
            {
                string filePath = dataRowView["FilePath"].ToString();

                if (File.Exists(filePath))
                {
                    try
                    {
                        Process.Start(new ProcessStartInfo(filePath) { UseShellExecute = true });
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Could not open file: {ex.Message}");
                    }
                }
                else
                {
                    MessageBox.Show("The file no longer exists in the Vault. It may have been manually deleted from the system folder.", "File Not Found", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        // NEW: Handles deleting the document from DB and deleting the physical file
        private void btnDeleteDocument_Click(object sender, RoutedEventArgs e)
        {
            var button = sender as Button;
            var dataRowView = button.DataContext as DataRowView;

            if (dataRowView != null)
            {
                string docId = dataRowView["DocID"].ToString();
                string docName = dataRowView["DocumentName"].ToString();
                string filePath = dataRowView["FilePath"].ToString();

                // Ask for confirmation before deleting
                MessageBoxResult result = MessageBox.Show($"Are you sure you want to permanently delete the document '{docName}'?\n\nThis cannot be undone.", "Confirm Delete", MessageBoxButton.YesNo, MessageBoxImage.Warning);

                if (result == MessageBoxResult.Yes)
                {
                    try
                    {
                        // 1. Delete from SQLite Database
                        using (var connection = new SqliteConnection(DatabaseHelper.GetConnectionString()))
                        {
                            connection.Open();
                            string query = "DELETE FROM EmployeeDocuments WHERE DocID = @DocID";

                            using (var command = new SqliteCommand(query, connection))
                            {
                                command.Parameters.AddWithValue("@DocID", docId);
                                command.ExecuteNonQuery();
                            }
                        }

                        // 2. Delete the physical file from the Vault to save space
                        if (File.Exists(filePath))
                        {
                            File.Delete(filePath);
                        }

                        // 3. Refresh the table
                        LoadDocuments();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Error deleting document: {ex.Message}", "Delete Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
            }
        }
    }
}