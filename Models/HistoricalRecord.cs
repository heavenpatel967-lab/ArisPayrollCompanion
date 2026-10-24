using System;

namespace ArisPayrollCompanion.Models
{
    public class HistoricalRecord
    {
        public DateTime SortDate { get; set; }
        public string MonthYear { get; set; } = string.Empty;
        public string Amount { get; set; } = string.Empty;
    }
}