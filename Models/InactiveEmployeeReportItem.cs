namespace ArisPayrollCompanion.Models
{
    public class InactiveEmployeeReportItem
    {
        public string EmployeeCode { get; set; } = "";
        public string EmployeeName { get; set; } = "";
        public string Department { get; set; } = "";
        public string Category { get; set; } = "";
        public string JoinDate { get; set; } = "";
        public string LeavingDate { get; set; } = "";
        public decimal PendingLoanAmount { get; set; }
    }
}