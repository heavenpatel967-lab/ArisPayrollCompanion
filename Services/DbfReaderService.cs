using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Linq;
using DotNetDBF;
using ArisPayrollCompanion.Models;

namespace ArisPayrollCompanion.Services
{
    public class DbfReaderService
    {
        public int GetRecordCount(string dbfFile)
        {
            if (!File.Exists(dbfFile)) return 0;
            try { using var r = new DBFReader(dbfFile); return r.RecordCount; } catch { return 0; }
        }

        public List<string> GetFieldNames(string dbfFile)
        {
            if (!File.Exists(dbfFile)) return new List<string>();
            try { using var r = new DBFReader(dbfFile); return r.Fields.Select(f => f.Name).ToList(); } catch { return new List<string>(); }
        }

        private List<Dictionary<string, string>> ReadDbfManuallySafely(string filePath)
        {
            var records = new List<Dictionary<string, string>>();
            if (!File.Exists(filePath)) return records;

            string tempFile = Path.GetTempFileName();
            try
            {
                using (var source = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var dest = new FileStream(tempFile, FileMode.Create, FileAccess.Write))
                {
                    source.CopyTo(dest);
                }

                using var stream = new FileStream(tempFile, FileMode.Open, FileAccess.Read);
                using var reader = new BinaryReader(stream, Encoding.Latin1);

                reader.ReadBytes(4);
                int fakeHeaderCount = reader.ReadInt32();
                int headerSize = reader.ReadInt16();
                int recordSize = reader.ReadInt16();
                reader.ReadBytes(20);

                var fields = new List<Tuple<string, int>>();

                while (reader.BaseStream.Position < headerSize - 1)
                {
                    byte firstByte = reader.ReadByte();
                    if (firstByte == 0x0D) break;
                    reader.BaseStream.Position--;

                    byte[] nameBytes = reader.ReadBytes(11);
                    string name = Encoding.ASCII.GetString(nameBytes).Replace("\0", "").Trim();

                    reader.ReadByte();
                    reader.ReadBytes(4);
                    int len = reader.ReadByte();
                    reader.ReadBytes(15);

                    if (!string.IsNullOrEmpty(name)) fields.Add(new Tuple<string, int>(name.ToUpper(), len));
                }

                reader.BaseStream.Position = headerSize;

                while (reader.BaseStream.Position + recordSize <= reader.BaseStream.Length)
                {
                    byte[] rowData = reader.ReadBytes(recordSize);
                    if (rowData.Length < recordSize) break;

                    if (rowData[0] == 0x2A) continue;

                    var rowDict = new Dictionary<string, string>();
                    int offset = 1;

                    foreach (var field in fields)
                    {
                        if (offset + field.Item2 <= rowData.Length)
                        {
                            string val = Encoding.Latin1.GetString(rowData, offset, field.Item2).Replace("\0", "").Trim();
                            rowDict[field.Item1] = val;
                        }
                        offset += field.Item2;
                    }
                    records.Add(rowDict);
                }
            }
            catch { }
            finally
            {
                try { if (File.Exists(tempFile)) File.Delete(tempFile); } catch { }
            }

            return records;
        }

        public List<Employee> ReadEmployees(string folderOrFilePath)
        {
            Dictionary<string, Employee> employeesMap = new();

            string folderPath = Directory.Exists(folderOrFilePath) ? folderOrFilePath : Path.GetDirectoryName(folderOrFilePath);
            if (!Directory.Exists(folderPath)) return new List<Employee>();

            string[] targetFiles = { "EMPMAS.DBF", "empmas.dbf", "mstsal.DBF", "mstsal.dbf", "monsum.DBF", "monsum.dbf" };

            foreach (string fileName in targetFiles)
            {
                string fullPath = Path.Combine(folderPath, fileName);
                if (!File.Exists(fullPath)) continue;

                var records = ReadDbfManuallySafely(fullPath);

                foreach (var record in records)
                {
                    string empCode = GetVal(record, "EMPCD");
                    string card = GetVal(record, "CARDNO");
                    if (string.IsNullOrEmpty(card)) card = GetVal(record, "COMP_CODE");

                    if (string.IsNullOrWhiteSpace(empCode)) empCode = card;
                    if (string.IsNullOrWhiteSpace(empCode)) continue;

                    string name = GetVal(record, "EMPNM");
                    if (string.IsNullOrEmpty(name)) name = GetVal(record, "EMP_NAME");
                    if (string.IsNullOrEmpty(name)) name = GetVal(record, "NAME");

                    string dept = GetVal(record, "DEPT");
                    string desig = GetVal(record, "DESIG");
                    string cat = GetVal(record, "CAT");

                    string leavingDt = GetVal(record, "LEAVING_DT").Replace("-", "").Replace("/", "").Replace("0", "").Trim();
                    string status = string.IsNullOrEmpty(leavingDt) ? "Y" : "N";

                    if (employeesMap.ContainsKey(empCode))
                    {
                        var existing = employeesMap[empCode];

                        if (string.IsNullOrWhiteSpace(existing.EmployeeName) || existing.EmployeeName == "Name Missing")
                            if (!string.IsNullOrWhiteSpace(name)) existing.EmployeeName = name;

                        if (string.IsNullOrWhiteSpace(existing.CardNumber) || existing.CardNumber == "00000001" || existing.CardNumber.Contains("0000000"))
                            if (!string.IsNullOrWhiteSpace(card)) existing.CardNumber = card;

                        if (string.IsNullOrWhiteSpace(existing.Department) && !string.IsNullOrWhiteSpace(dept)) existing.Department = dept;
                        if (string.IsNullOrWhiteSpace(existing.Designation) && !string.IsNullOrWhiteSpace(desig)) existing.Designation = desig;
                        if (string.IsNullOrWhiteSpace(existing.Category) && !string.IsNullOrWhiteSpace(cat)) existing.Category = cat;
                        if (existing.Status == "Y" && status == "N") existing.Status = "N";
                    }
                    else
                    {
                        employeesMap[empCode] = new Employee
                        {
                            EmployeeCode = empCode,
                            EmployeeName = string.IsNullOrWhiteSpace(name) ? "Name Missing" : name,
                            CardNumber = string.IsNullOrWhiteSpace(card) ? "" : card,
                            Department = dept,
                            Designation = desig,
                            Category = cat,
                            Status = status
                        };
                    }
                }
            }

            foreach (var emp in employeesMap.Values)
            {
                if (!string.IsNullOrWhiteSpace(emp.CardNumber) && emp.CardNumber.Length < 8)
                {
                    emp.CardNumber = emp.CardNumber.PadLeft(8, '0');
                }
            }

            return employeesMap.Values.OrderByDescending(e => int.TryParse(e.EmployeeCode, out int n) ? n : 0).ToList();
        }

        private string GetVal(Dictionary<string, string> record, string fieldName)
        {
            if (record.ContainsKey(fieldName)) return record[fieldName];
            return "";
        }

        public List<DailyAttendance> GetDailyAttendance(string exactFilePath, string targetDate)
        {
            var records = new List<DailyAttendance>();

            if (!File.Exists(exactFilePath)) return records;

            string tempFile = Path.GetTempFileName();
            try
            {
                using (var source = new FileStream(exactFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var dest = new FileStream(tempFile, FileMode.Create, FileAccess.Write))
                {
                    source.CopyTo(dest);
                }

                using var stream = new FileStream(tempFile, FileMode.Open, FileAccess.Read);
                using var reader = new BinaryReader(stream, Encoding.Latin1);

                reader.ReadBytes(4);
                reader.ReadInt32();
                int headerSize = reader.ReadInt16();
                int recordSize = reader.ReadInt16();
                reader.ReadBytes(20);

                var fields = new List<Tuple<string, int>>();
                while (reader.BaseStream.Position < headerSize - 1)
                {
                    byte firstByte = reader.ReadByte();
                    if (firstByte == 0x0D) break;
                    reader.BaseStream.Position--;

                    byte[] nameBytes = reader.ReadBytes(11);
                    string name = Encoding.ASCII.GetString(nameBytes).Replace("\0", "").Trim().ToUpper();

                    reader.ReadByte(); reader.ReadBytes(4);
                    int len = reader.ReadByte();
                    reader.ReadBytes(15);

                    if (!string.IsNullOrEmpty(name)) fields.Add(new Tuple<string, int>(name, len));
                }

                reader.BaseStream.Position = headerSize;

                int empCdIdx = fields.FindIndex(f => f.Item1 == "EMPCD");
                int dateIdx = fields.FindIndex(f => f.Item1 == "DATE");
                int in1Idx = fields.FindIndex(f => f.Item1 == "IN1");
                int absIdx = fields.FindIndex(f => f.Item1 == "ABS");
                int off1Idx = fields.FindIndex(f => f.Item1 == "OFF1");

                if (empCdIdx == -1 || dateIdx == -1) return records;

                while (reader.BaseStream.Position + recordSize <= reader.BaseStream.Length)
                {
                    byte[] rowData = reader.ReadBytes(recordSize);
                    if (rowData.Length < recordSize) break;
                    if (rowData[0] == 0x2A) continue; // Skip deleted rows

                    string rowDate = ExtractByteString(rowData, fields, dateIdx);

                    if (rowDate == targetDate)
                    {
                        string in1Str = ExtractByteString(rowData, fields, in1Idx);
                        string absStr = ExtractByteString(rowData, fields, absIdx);

                        double.TryParse(in1Str, out double inTime);
                        double.TryParse(absStr, out double absValue);

                        records.Add(new DailyAttendance
                        {
                            EmployeeCode = ExtractByteString(rowData, fields, empCdIdx),
                            Date = rowDate,
                            InTime = inTime,
                            IsAbsent = absValue >= 1.0,
                            WeekOffDay = ExtractByteString(rowData, fields, off1Idx)
                        });
                    }
                }
            }
            catch { }
            finally
            {
                try { if (File.Exists(tempFile)) File.Delete(tempFile); } catch { }
            }

            return records;
        }

        // 🔥 NEW: Extracts punch history for a SINGLE employee across a DATE RANGE
        public List<EmployeeDailyPunch> GetEmployeeAttendanceHistory(string exactFilePath, string empCode, DateTime startDate, DateTime endDate)
        {
            var records = new List<EmployeeDailyPunch>();
            if (!File.Exists(exactFilePath)) return records;

            string tempFile = Path.GetTempFileName();
            try
            {
                using (var source = new FileStream(exactFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var dest = new FileStream(tempFile, FileMode.Create, FileAccess.Write))
                {
                    source.CopyTo(dest);
                }

                using var stream = new FileStream(tempFile, FileMode.Open, FileAccess.Read);
                using var reader = new BinaryReader(stream, Encoding.Latin1);

                reader.ReadBytes(4);
                reader.ReadInt32();
                int headerSize = reader.ReadInt16();
                int recordSize = reader.ReadInt16();
                reader.ReadBytes(20);

                var fields = new List<Tuple<string, int>>();
                while (reader.BaseStream.Position < headerSize - 1)
                {
                    byte firstByte = reader.ReadByte();
                    if (firstByte == 0x0D) break;
                    reader.BaseStream.Position--;

                    byte[] nameBytes = reader.ReadBytes(11);
                    string name = Encoding.ASCII.GetString(nameBytes).Replace("\0", "").Trim().ToUpper();

                    reader.ReadByte(); reader.ReadBytes(4);
                    int len = reader.ReadByte();
                    reader.ReadBytes(15);

                    if (!string.IsNullOrEmpty(name)) fields.Add(new Tuple<string, int>(name, len));
                }

                reader.BaseStream.Position = headerSize;

                int empCdIdx = fields.FindIndex(f => f.Item1 == "EMPCD");
                int dateIdx = fields.FindIndex(f => f.Item1 == "DATE");
                int in1Idx = fields.FindIndex(f => f.Item1 == "IN1");
                int out1Idx = fields.FindIndex(f => f.Item1 == "OUT1");
                int absIdx = fields.FindIndex(f => f.Item1 == "ABS");
                int off1Idx = fields.FindIndex(f => f.Item1 == "OFF1");

                if (empCdIdx == -1 || dateIdx == -1) return records;

                string[] formats = { "yyyyMMdd", "dd-MM-yyyy", "MM/dd/yyyy", "yyyy-MM-dd", "dd/MM/yyyy" };

                // Handle employee code padding to match raw DBF data
                string targetEmp1 = empCode.Trim();
                string targetEmp2 = empCode.PadLeft(4, '0');

                while (reader.BaseStream.Position + recordSize <= reader.BaseStream.Length)
                {
                    byte[] rowData = reader.ReadBytes(recordSize);
                    if (rowData.Length < recordSize) break;
                    if (rowData[0] == 0x2A) continue;

                    string rowEmp = ExtractByteString(rowData, fields, empCdIdx);

                    if (rowEmp == targetEmp1 || rowEmp == targetEmp2)
                    {
                        string rowDateStr = ExtractByteString(rowData, fields, dateIdx);
                        if (DateTime.TryParseExact(rowDateStr, formats, null, System.Globalization.DateTimeStyles.None, out DateTime rowDate))
                        {
                            if (rowDate.Date >= startDate.Date && rowDate.Date <= endDate.Date)
                            {
                                string in1Str = ExtractByteString(rowData, fields, in1Idx);
                                string out1Str = out1Idx != -1 ? ExtractByteString(rowData, fields, out1Idx) : "";
                                string absStr = ExtractByteString(rowData, fields, absIdx);
                                string wOff = ExtractByteString(rowData, fields, off1Idx).ToUpper();

                                double.TryParse(in1Str, out double inTime);
                                double.TryParse(absStr, out double absValue);

                                string status = "Present";
                                string color = "#27AE60"; // Green

                                if (absValue >= 1.0)
                                {
                                    if (rowDate.DayOfWeek == DayOfWeek.Wednesday && wOff.Contains("WED"))
                                    {
                                        status = "Week Off";
                                        color = "#8E44AD"; // Purple
                                    }
                                    else
                                    {
                                        status = "Absent";
                                        color = "#E74C3C"; // Red
                                    }
                                }
                                else if (inTime > 8.32)
                                {
                                    status = "Late";
                                    color = "#F39C12"; // Orange
                                }

                                records.Add(new EmployeeDailyPunch
                                {
                                    DateStr = rowDate.ToString("dd-MMM-yyyy"),
                                    SortDate = rowDate,
                                    InTime = inTime > 0 ? inTime.ToString("0.00") : "-",
                                    OutTime = string.IsNullOrWhiteSpace(out1Str) || out1Str == "0.00" ? "-" : double.Parse(out1Str).ToString("0.00"),
                                    Status = status,
                                    StatusColor = color
                                });
                            }
                        }
                    }
                }
            }
            catch { }
            finally { try { if (File.Exists(tempFile)) File.Delete(tempFile); } catch { } }

            return records.OrderBy(r => r.SortDate).ToList(); // Sorts oldest to newest
        }

        private string ExtractByteString(byte[] rowData, List<Tuple<string, int>> fields, int index)
        {
            if (index == -1) return "";
            int offset = 1;
            for (int i = 0; i < index; i++) offset += fields[i].Item2;
            if (offset + fields[index].Item2 <= rowData.Length)
                return Encoding.Latin1.GetString(rowData, offset, fields[index].Item2).Trim();
            return "";
        }
    }
}