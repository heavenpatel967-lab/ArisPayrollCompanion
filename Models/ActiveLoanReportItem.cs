namespace ArisPayrollCompanion.Models
{
    public class ActiveLoanReportItem
    {
        public string CardNumber { get; set; } = "";
        public string EmployeeCode { get; set; } = "";
        public string EmployeeName { get; set; } = "";
        public string Department { get; set; } = "";
        public string Category { get; set; } = "";
        public decimal PendingLoanAmount { get; set; }
    }
}