using System.IO;
using ArisPayrollCompanion.Repositories;

namespace ArisPayrollCompanion.Services
{
    public class SyncService
    {
        private readonly DbfReaderService _dbfReader;
        private readonly EmployeeRepository _employeeRepository;

        public SyncService()
        {
            _dbfReader = new DbfReaderService();
            _employeeRepository = new EmployeeRepository();
        }

        public int SyncEmployees(string dbfFolderOrFile)
        {
            // Ensure we pass the directory folder path to the reader
            string folderPath = Directory.Exists(dbfFolderOrFile)
                ? dbfFolderOrFile
                : Path.GetDirectoryName(dbfFolderOrFile) ?? "";

            string monsumPath = Path.Combine(folderPath, "monsum.DBF");
            if (!File.Exists(monsumPath)) monsumPath = Path.Combine(folderPath, "monsum.dbf");
            if (!File.Exists(monsumPath)) monsumPath = Path.Combine(folderPath, dbfFolderOrFile); // Fallback to whatever file was passed

            var employees = _dbfReader.ReadEmployees(monsumPath);

            _employeeRepository.DeleteAllEmployees();
            _employeeRepository.SaveEmployees(employees);

            return employees.Count;
        }

        public int GetEmployeeCount()
        {
            return _employeeRepository.GetEmployeeCount();
        }
    }
}