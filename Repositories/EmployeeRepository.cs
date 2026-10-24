using System;
using System.Collections.Generic;
using Microsoft.Data.Sqlite;
using ArisPayrollCompanion.Models;
using ArisPayrollCompanion.Helpers;

namespace ArisPayrollCompanion.Repositories
{
    public class EmployeeRepository
    {
        private string GetConnectionString() => DatabaseHelper.GetConnectionString();

        public EmployeeRepository()
        {
            CreateTable();
        }

        private void CreateTable()
        {
            using var connection = new SqliteConnection(GetConnectionString());
            connection.Open();

            string sql = @"
CREATE TABLE IF NOT EXISTS Employees
(
    EmployeeCode TEXT PRIMARY KEY,
    CardNumber TEXT,
    EmployeeName TEXT,
    Department TEXT,
    Designation TEXT,
    Category TEXT,
    Status TEXT
);";

            using var command = new SqliteCommand(sql, connection);
            command.ExecuteNonQuery();
        }

        public void DeleteAllEmployees()
        {
            using var connection = new SqliteConnection(GetConnectionString());
            connection.Open();
            using var command = new SqliteCommand("DELETE FROM Employees", connection);
            command.ExecuteNonQuery();
        }

        public void SaveEmployees(List<Employee> employees)
        {
            using var connection = new SqliteConnection(GetConnectionString());
            connection.Open();
            using var transaction = connection.BeginTransaction();

            foreach (var employee in employees)
            {
                using var command = new SqliteCommand(@"
INSERT OR REPLACE INTO Employees
(EmployeeCode, CardNumber, EmployeeName, Department, Designation, Category, Status)
VALUES
(@EmployeeCode, @CardNumber, @EmployeeName, @Department, @Designation, @Category, @Status);", connection);

                command.Transaction = transaction;
                command.Parameters.AddWithValue("@EmployeeCode", employee.EmployeeCode);
                command.Parameters.AddWithValue("@CardNumber", employee.CardNumber ?? "");
                command.Parameters.AddWithValue("@EmployeeName", employee.EmployeeName ?? "");
                command.Parameters.AddWithValue("@Department", employee.Department ?? "");
                command.Parameters.AddWithValue("@Designation", employee.Designation ?? "");
                command.Parameters.AddWithValue("@Category", employee.Category ?? "");
                command.Parameters.AddWithValue("@Status", employee.Status ?? "Y");

                command.ExecuteNonQuery();
            }
            transaction.Commit();
        }

        public void UpdateEmployee(Employee employee)
        {
            using var connection = new SqliteConnection(GetConnectionString());
            connection.Open();

            string sql = @"
                UPDATE Employees 
                SET CardNumber = @CardNumber, 
                    EmployeeName = @EmployeeName, 
                    Department = @Department, 
                    Designation = @Designation, 
                    Category = @Category, 
                    Status = @Status 
                WHERE EmployeeCode = @EmployeeCode";

            using var command = new SqliteCommand(sql, connection);
            command.Parameters.AddWithValue("@EmployeeCode", employee.EmployeeCode);
            command.Parameters.AddWithValue("@CardNumber", employee.CardNumber ?? "");
            command.Parameters.AddWithValue("@EmployeeName", employee.EmployeeName ?? "");
            command.Parameters.AddWithValue("@Department", employee.Department ?? "");
            command.Parameters.AddWithValue("@Designation", employee.Designation ?? "");
            command.Parameters.AddWithValue("@Category", employee.Category ?? "");
            command.Parameters.AddWithValue("@Status", employee.Status ?? "Y");

            command.ExecuteNonQuery();
        }

        public List<string> GetDistinctValues(string columnName, string defaultOption)
        {
            List<string> values = new() { defaultOption };

            using var connection = new SqliteConnection(GetConnectionString());
            connection.Open();

            string sql = $"SELECT DISTINCT {columnName} FROM Employees WHERE {columnName} != '' AND {columnName} IS NOT NULL ORDER BY {columnName}";

            using var command = new SqliteCommand(sql, connection);
            using var reader = command.ExecuteReader();

            while (reader.Read())
            {
                values.Add(reader[0].ToString() ?? "");
            }

            return values;
        }

        public List<Employee> GetAllEmployees()
        {
            return SearchEmployees("", "All Departments", "All Designations", "All Categories", "All Statuses");
        }

        public List<Employee> SearchEmployees(string searchTerm, string department, string designation, string category, string status)
        {
            List<Employee> employees = new();
            using var connection = new SqliteConnection(GetConnectionString());
            connection.Open();

            string sql = @"SELECT * FROM Employees WHERE 1=1";

            if (!string.IsNullOrWhiteSpace(searchTerm))
            {
                sql += @" AND (EmployeeName LIKE @search OR EmployeeCode LIKE @search OR CardNumber LIKE @search)";
            }

            if (department != "All Departments") sql += " AND Department = @department";
            if (designation != "All Designations") sql += " AND Designation = @designation";
            if (category != "All Categories") sql += " AND Category = @category";
            if (status != "All Statuses") sql += " AND Status = @status";

            sql += " ORDER BY EmployeeName";

            using var command = new SqliteCommand(sql, connection);

            if (!string.IsNullOrWhiteSpace(searchTerm))
            {
                command.Parameters.AddWithValue("@search", $"%{searchTerm}%");
            }

            if (department != "All Departments") command.Parameters.AddWithValue("@department", department);
            if (designation != "All Designations") command.Parameters.AddWithValue("@designation", designation);
            if (category != "All Categories") command.Parameters.AddWithValue("@category", category);
            if (status != "All Statuses") command.Parameters.AddWithValue("@status", status);

            using var reader = command.ExecuteReader();

            while (reader.Read())
            {
                employees.Add(new Employee
                {
                    EmployeeCode = reader["EmployeeCode"].ToString() ?? "",
                    CardNumber = reader["CardNumber"].ToString() ?? "",
                    EmployeeName = reader["EmployeeName"].ToString() ?? "",
                    Department = reader["Department"].ToString() ?? "",
                    Designation = reader["Designation"].ToString() ?? "",
                    Category = reader["Category"].ToString() ?? "",
                    Status = reader["Status"].ToString() ?? "Y"
                });
            }

            return employees;
        }

        public int GetEmployeeCount()
        {
            using var connection = new SqliteConnection(GetConnectionString());
            connection.Open();
            using var command = new SqliteCommand("SELECT COUNT(*) FROM Employees", connection);
            return Convert.ToInt32(command.ExecuteScalar());
        }
    }
}