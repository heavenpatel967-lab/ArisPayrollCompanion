namespace ArisPayrollCompanion.Models
{
    public class Employee
    {
        public string EmployeeCode { get; set; } = "";
        public string CardNumber { get; set; } = "";
        public string EmployeeName { get; set; } = "";
        public string Department { get; set; } = "";
        public string Designation { get; set; } = "";
        public string Category { get; set; } = "";
        public string Status { get; set; } = "";
    }

    public class DailyAttendance
    {
        public string EmployeeCode { get; set; } = "";
        public string Date { get; set; } = "";
        public double InTime { get; set; }
        public bool IsAbsent { get; set; }
        public string WeekOffDay { get; set; } = "";
    }

    // 🔥 UPDATED: Added CardNumber to the report model
    public class ReportResult
    {
        public string EmployeeCode { get; set; } = "";
        public string CardNumber { get; set; } = "";
        public string EmployeeName { get; set; } = "";
        public string Department { get; set; } = "";
        public string PunchTimeStr { get; set; } = "";
        public string Status { get; set; } = "";
    }

    // 🔥 NEW: Holds the row data for the Employee History Report
    public class EmployeeDailyPunch
    {
        public string DateStr { get; set; } = "";
        public System.DateTime SortDate { get; set; }
        public string InTime { get; set; } = "";
        public string OutTime { get; set; } = "";
        public string Status { get; set; } = "";
        public string StatusColor { get; set; } = "";
    }
}