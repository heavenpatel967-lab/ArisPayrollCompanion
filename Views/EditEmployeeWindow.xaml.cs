using System;
using System.IO;
using System.Windows;
using System.Windows.Media.Imaging;
using Microsoft.Win32;
using Microsoft.Data.Sqlite;
using ArisPayrollCompanion.Helpers;

namespace ArisPayrollCompanion.Views
{
    public partial class EditEmployeeWindow : Window
    {
        private string _photoFilePath = "";

        public EditEmployeeWindow(string employeeCode, string dbfName)
        {
            InitializeComponent();

            txtEmpCode.Text = employeeCode;
            txtDbfName.Text = dbfName;

            LoadEmployeeProfile(employeeCode);
        }

        private void LoadEmployeeProfile(string empCode)
        {
            try
            {
                using (var connection = new SqliteConnection(DatabaseHelper.GetConnectionString()))
                {
                    connection.Open();
                    string query = "SELECT * FROM EmployeeProfiles WHERE EmployeeCode = @Code";

                    using (var command = new SqliteCommand(query, connection))
                    {
                        command.Parameters.AddWithValue("@Code", empCode);

                        using (var reader = command.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                txtOriginalName.Text = reader["OriginalName"].ToString();
                                txtAadharName.Text = reader["AadharName"].ToString();
                                txtAadharNumber.Text = reader["AadharNumber"].ToString();
                                txtPhone.Text = reader["PhoneNumber"].ToString();
                                txtNativePlace.Text = reader["NativePlace"].ToString();
                                txtVillage.Text = reader["Village"].ToString();
                                txtRemarks.Text = reader["Remarks"].ToString();
                                txtDuties.Text = reader["DutiesAssigned"].ToString();

                                string savedPhotoPath = reader["PhotoFilePath"].ToString();
                                if (!string.IsNullOrEmpty(savedPhotoPath) && File.Exists(savedPhotoPath))
                                {
                                    _photoFilePath = savedPhotoPath;
                                    LoadImageToUI(_photoFilePath);
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading profile: {ex.Message}");
            }
        }

        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // ==========================================
                // THE PHOTO VAULT SYSTEM
                // ==========================================
                string finalPhotoPath = _photoFilePath;

                // Check if the photo is new (doesn't already exist in the Vault)
                if (!string.IsNullOrWhiteSpace(_photoFilePath) && !_photoFilePath.Contains("CompanionVault"))
                {
                    string baseAppFolder = AppDomain.CurrentDomain.BaseDirectory;
                    string vaultFolder = Path.Combine(baseAppFolder, "CompanionVault", "ProfilePhotos");

                    if (!Directory.Exists(vaultFolder))
                    {
                        Directory.CreateDirectory(vaultFolder);
                    }

                    // Create unique file name: EmployeeCode + Timestamp
                    string ext = Path.GetExtension(_photoFilePath);
                    string uniqueFileName = $"{txtEmpCode.Text}_Profile_{DateTime.Now:yyyyMMdd_HHmmss}{ext}";
                    finalPhotoPath = Path.Combine(vaultFolder, uniqueFileName);

                    // Securely copy the photo to the Vault
                    File.Copy(_photoFilePath, finalPhotoPath, true);
                }
                // ==========================================

                using (var connection = new SqliteConnection(DatabaseHelper.GetConnectionString()))
                {
                    connection.Open();

                    string query = @"
                        INSERT OR REPLACE INTO EmployeeProfiles 
                        (EmployeeCode, OriginalName, AadharName, AadharNumber, PhoneNumber, NativePlace, Village, PhotoFilePath, Remarks, DutiesAssigned) 
                        VALUES 
                        (@Code, @OrigName, @AadharName, @AadharNum, @Phone, @Native, @Village, @Photo, @Remarks, @Duties)";

                    using (var command = new SqliteCommand(query, connection))
                    {
                        command.Parameters.AddWithValue("@Code", txtEmpCode.Text);
                        command.Parameters.AddWithValue("@OrigName", txtOriginalName.Text);
                        command.Parameters.AddWithValue("@AadharName", txtAadharName.Text);
                        command.Parameters.AddWithValue("@AadharNum", txtAadharNumber.Text);
                        command.Parameters.AddWithValue("@Phone", txtPhone.Text);
                        command.Parameters.AddWithValue("@Native", txtNativePlace.Text);
                        command.Parameters.AddWithValue("@Village", txtVillage.Text);

                        // Save the secure Vault path
                        command.Parameters.AddWithValue("@Photo", finalPhotoPath);

                        command.Parameters.AddWithValue("@Remarks", txtRemarks.Text);
                        command.Parameters.AddWithValue("@Duties", txtDuties.Text);

                        command.ExecuteNonQuery();
                    }
                }

                MessageBox.Show("Profile saved successfully!", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
                this.DialogResult = true;
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error saving profile: {ex.Message}", "Database Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void btnBrowsePhoto_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Title = "Select Profile Photo",
                Filter = "Image Files|*.jpg;*.jpeg;*.png;*.bmp"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                _photoFilePath = openFileDialog.FileName;
                LoadImageToUI(_photoFilePath);
            }
        }

        // NEW: Upgraded to securely delete the file and update the database
        private void btnClearPhoto_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(_photoFilePath)) return;

            MessageBoxResult result = MessageBox.Show("Are you sure you want to permanently delete this profile photo?", "Confirm Delete", MessageBoxButton.YesNo, MessageBoxImage.Warning);

            if (result == MessageBoxResult.Yes)
            {
                try
                {
                    // 1. Delete physical file if it is stored in the Vault
                    if (_photoFilePath.Contains("CompanionVault") && File.Exists(_photoFilePath))
                    {
                        File.Delete(_photoFilePath);
                    }

                    // 2. Clear from SQLite Database
                    using (var connection = new SqliteConnection(DatabaseHelper.GetConnectionString()))
                    {
                        connection.Open();
                        string query = "UPDATE EmployeeProfiles SET PhotoFilePath = '' WHERE EmployeeCode = @Code";
                        using (var command = new SqliteCommand(query, connection))
                        {
                            command.Parameters.AddWithValue("@Code", txtEmpCode.Text);
                            command.ExecuteNonQuery();
                        }
                    }

                    // 3. Reset the UI
                    _photoFilePath = "";
                    imgProfilePhoto.Source = null;
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error deleting photo: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void LoadImageToUI(string filePath)
        {
            try
            {
                BitmapImage bitmap = new BitmapImage();
                bitmap.BeginInit();
                bitmap.UriSource = new Uri(filePath);
                bitmap.CacheOption = BitmapCacheOption.OnLoad; // Crucial: Prevents Windows from locking the file so we can delete it later
                bitmap.EndInit();

                imgProfilePhoto.Source = bitmap;
            }
            catch
            {
                MessageBox.Show("Could not load the selected image.", "Image Error");
            }
        }

        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}