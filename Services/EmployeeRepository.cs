using ClosedXML.Excel;
using BiometricAttendanceProcessor.Configuration;
using BiometricAttendanceProcessor.Models;

namespace BiometricAttendanceProcessor.Services;

/// <summary>
/// Reads employee master data from the "CriteriaSheet" worksheet.
/// </summary>
public interface IEmployeeRepository
{
    IReadOnlyList<Employee> GetAllEmployees();
}

/// <summary>
/// DESIGN PATTERN: REPOSITORY.
///
/// A Repository's job is to hide WHERE data comes from behind a simple,
/// storage-agnostic interface (IEmployeeRepository above). Every other class in
/// this program depends only on IEmployeeRepository, never on EmployeeRepository
/// or on ClosedXML directly. That means:
///
///   - If employee master data ever moves from an Excel sheet to a SQL database
///     or a REST API, you write ONE new class (e.g. SqlEmployeeRepository) that
///     implements IEmployeeRepository, and swap it in at the composition root
///     (Program.cs). AttendanceProcessor, AttendanceAnalyzer, and everything
///     downstream needs ZERO changes.
///   - You can write unit tests for AttendanceProcessor using a fake, in-memory
///     IEmployeeRepository, with no Excel file involved at all.
///
/// This is the same reasoning the VBA "CEmployeeRepository" class was reaching
/// for, but VBA has no interfaces, so that reasoning could never be fully
/// realized there - every caller was still implicitly coupled to "a Worksheet".
/// </summary>
public class EmployeeRepository : IEmployeeRepository
{
    private readonly IXLWorksheet _sheet;

    public EmployeeRepository(IXLWorksheet criteriaSheet)
    {
        _sheet = criteriaSheet;
    }

    public IReadOnlyList<Employee> GetAllEmployees()
    {
        var employees = new List<Employee>();
        int lastRow = _sheet.LastRowUsed()?.RowNumber() ?? 1;

        // Row 1 is the header row, so data starts at row 2.
        for (int row = 2; row <= lastRow; row++)
        {
            employees.Add(ReadEmployeeRow(row));
        }

        return employees;
    }

    private Employee ReadEmployeeRow(int row)
    {
        var rawId = CleanString(_sheet.Cell(row, (int)CriteriaColumn.EmployeeId).GetString());

        var employee = new Employee
        {
            EmployeeId = string.IsNullOrWhiteSpace(rawId) ? "No EmpID" : rawId,
            Name = CleanString(_sheet.Cell(row, (int)CriteriaColumn.EmployeeName).GetString()),
            Type = ContractTypeParser.Parse(_sheet.Cell(row, (int)CriteriaColumn.ContractType).GetString())
        };

        // Matches the original VBA behaviour: Location/JobTitle are only read
        // (and therefore only meaningful) when the employee has a real ID.
        if (employee.HasValidId)
        {
            employee.Location = CleanString(_sheet.Cell(row, (int)CriteriaColumn.Location).GetString());
            employee.JobTitle = CleanString(_sheet.Cell(row, (int)CriteriaColumn.JobTitle).GetString());
        }

        var startDate = _sheet.Cell(row, (int)CriteriaColumn.StartDate).GetDateTime();
        var endDate = _sheet.Cell(row, (int)CriteriaColumn.EndDate).GetDateTime();

        // Consultants: working days = NETWORKDAYS equivalent (Mon-Fri count).
        // Contractors: working days = every calendar day inclusive.
        employee.WorkingDaysInMonth = employee.Type == ContractType.Consultant
            ? CountWeekdays(startDate, endDate)
            : (endDate.Date - startDate.Date).Days + 1;

        return employee;
    }

    /// <summary>
    /// Equivalent of Excel's NETWORKDAYS(start, end) with no holiday exclusion
    /// list supplied - matching the original VBA call
    /// WorksheetFunction.NetworkDays(startDate, endDate), which likewise passed
    /// no holiday argument.
    /// </summary>
    private static long CountWeekdays(DateTime start, DateTime end)
    {
        long count = 0;
        for (var d = start.Date; d <= end.Date; d = d.AddDays(1))
        {
            if (d.DayOfWeek is not (DayOfWeek.Saturday or DayOfWeek.Sunday))
                count++;
        }
        return count;
    }

    /// <summary>
    /// Removes non-breaking spaces and trims/collapses whitespace, matching the
    /// VBA CleanExcelString utility function.
    /// </summary>
    private static string CleanString(string? input)
    {
        if (string.IsNullOrEmpty(input)) return string.Empty;
        var cleaned = input.Replace('\u00A0', ' ');
        return System.Text.RegularExpressions.Regex.Replace(cleaned, @"\s+", " ").Trim();
    }
}
