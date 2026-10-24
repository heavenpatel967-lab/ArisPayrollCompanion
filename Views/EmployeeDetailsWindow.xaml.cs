using System.Windows;
using ArisPayrollCompanion.Models;

namespace ArisPayrollCompanion.Views
{
    public partial class EmployeeDetailsWindow : Window
    {
        private readonly Employee _employee;

        public EmployeeDetailsWindow(Employee employee)
        {
            InitializeComponent();
            _employee = employee;
            DisplayEmployeeDetails();
        }

        private void DisplayEmployeeDetails()
        {
            if (_employee != null)
            {
                txtName.Text = _employee.EmployeeName;
                txtCode.Text = _employee.EmployeeCode;
                txtCard.Text = _employee.CardNumber;

                // Translated Codes to Names
                txtDept.Text = GetDepartmentName(_employee.Department);
                txtCategory.Text = GetCategoryName(_employee.Category);

                txtDesig.Text = _employee.Designation;
                txtStatus.Text = _employee.Status;
            }
        }

        private void btnLoanLedger_Click(object sender, RoutedEventArgs e)
        {
            LoanReportWindow loanWindow = new LoanReportWindow(_employee);
            loanWindow.Owner = this;
            loanWindow.ShowDialog();
        }

        private void btnAdvance_Click(object sender, RoutedEventArgs e)
        {
            AdvanceReportWindow advanceWindow = new AdvanceReportWindow(_employee);
            advanceWindow.Owner = this;
            advanceWindow.ShowDialog();
        }

        private void btnLeaves_Click(object sender, RoutedEventArgs e)
        {
            LeaveTrackerWindow leaveWindow = new LeaveTrackerWindow(_employee.EmployeeCode, _employee.EmployeeName);
            leaveWindow.Owner = this;
            leaveWindow.ShowDialog();
        }

        private void btnDocuments_Click(object sender, RoutedEventArgs e)
        {
            EmployeeDocumentsWindow docWindow = new EmployeeDocumentsWindow(_employee.EmployeeCode, _employee.EmployeeName);
            docWindow.Owner = this;
            docWindow.ShowDialog();
        }

        private void btnSummary_Click(object sender, RoutedEventArgs e)
        {
            EmployeeSummaryWindow summaryWindow = new EmployeeSummaryWindow(_employee);
            summaryWindow.Owner = this;
            summaryWindow.ShowDialog();
        }

        private void btnEdit_Click(object sender, RoutedEventArgs e)
        {
            EditEmployeeWindow editWindow = new EditEmployeeWindow(_employee.EmployeeCode, _employee.EmployeeName);
            editWindow.Owner = this;

            if (editWindow.ShowDialog() == true)
            {
                DisplayEmployeeDetails();
            }
        }

        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        // ==========================================
        // MASTER LIST TRANSLATORS
        // ==========================================

        private string GetDepartmentName(string code)
        {
            if (string.IsNullOrWhiteSpace(code)) return "";

            switch (code.Trim().TrimStart('0'))
            {
                case "1": return "OFFICE";
                case "2": return "CASTING";
                case "3": return "MACHINE";
                case "4": return "BUFF & AMRI";
                case "5": return "CHROM";
                case "6": return "FITTING / OTHER";
                case "7": return "OTHER NEW";
                default: return code; // Returns original code if not found
            }
        }

        private string GetCategoryName(string code)
        {
            if (string.IsNullOrWhiteSpace(code)) return "";

            switch (code.Trim().TrimStart('0'))
            {
                case "1": return "STAFF";
                case "2": return "CASTING_10";
                case "3": return "WORKER";
                case "4": return "MONTHLY_8";
                case "5": return "GENERAL / PIC RATE";
                case "6": return "GENERAL ODD";
                case "7": return "MONTHLY_10";
                case "8": return "TRANSPORT";
                case "10": return "PLUMBER";
                case "11": return "PIC RATE GENERAL";
                case "12": return "WORKER_9";
                default: return code; // Returns original code if not found
            }
        }
    }
}