using ClosedXML.Excel;
using BiometricAttendanceProcessor.Analysis;
using BiometricAttendanceProcessor.Configuration;
using BiometricAttendanceProcessor.Models;
using BiometricAttendanceProcessor.Reporting;
using BiometricAttendanceProcessor.Services;

namespace BiometricAttendanceProcessor.Processing;

/// <summary>
/// DESIGN PATTERN: FACADE.
///
/// AttendanceProcessor hides the coordination of SIX collaborating classes
/// (PayrollCalendar, IEmployeeRepository, IBiometricDataService,
/// AttendanceAnalyzer, AttendanceReportGenerator, FileSystemService) behind one
/// simple public method: ProcessAll(). Program.cs (or, in the future, a test
/// project, a web API controller, or a scheduled background job) never needs to
/// know these six classes exist or how they collaborate - it just constructs an
/// AttendanceProcessor and calls ProcessAll().
///
/// This is the direct equivalent of the VBA "CAttendanceProcessor" class, which
/// described itself as implementing the Facade pattern in its own header
/// comment - that intent is fully realized here, backed by real interfaces
/// (IEmployeeRepository, IBiometricDataService) that VBA has no equivalent for.
/// </summary>
public class AttendanceProcessor
{
    private readonly IEmployeeRepository _employeeRepository;
    private readonly IBiometricDataService _biometricService;
    private readonly AttendanceReportGenerator _reportGenerator;
    private readonly FileSystemService _fileService;
    private readonly PayrollCalendar _calendar;
    private readonly bool _saveExtractedFilesPerEmployee;

    private List<BiometricRecord>? _allBiometricRecords;
    private int _processedCount;
    private int _notFoundCount;

    public AttendanceProcessor(
        IEmployeeRepository employeeRepository,
        IBiometricDataService biometricService,
        AttendanceReportGenerator reportGenerator,
        FileSystemService fileService,
        PayrollCalendar calendar,
        bool saveExtractedFilesPerEmployee = true)
    {
        _employeeRepository = employeeRepository;
        _biometricService = biometricService;
        _reportGenerator = reportGenerator;
        _fileService = fileService;
        _calendar = calendar;
        _saveExtractedFilesPerEmployee = saveExtractedFilesPerEmployee;
    }

    /// <summary>Processes every employee found in CriteriaSheet against the biometric DataSheet.</summary>
    public void ProcessAll()
    {
        _allBiometricRecords = _biometricService.LoadAll();
        var employees = _employeeRepository.GetAllEmployees();

        foreach (var employee in employees)
        {
            ProcessSingleEmployee(employee);
        }

        Console.WriteLine($"Successfully analyzed {_processedCount} biometric attendance record(s).");
        if (_notFoundCount > 0)
        {
            Console.WriteLine($"{_notFoundCount} employee(s) had no valid ID or no matching biometric data.");
        }
    }

    private void ProcessSingleEmployee(Employee employee)
    {
        if (!employee.HasValidId)
        {
            _reportGenerator.WriteEmployeeNotFound(employee);
            _notFoundCount++;
            return;
        }

        int startIndex = _biometricService.FindFirstRowIndexForEmployee(_allBiometricRecords!, employee.EmployeeId);
        if (startIndex == -1)
        {
            _reportGenerator.WriteEmployeeNotFound(employee);
            _notFoundCount++;
            return;
        }

        var firstDate = _biometricService.GetFirstDateForEmployee(_allBiometricRecords!, employee.EmployeeId, startIndex);
        bool isIrregular = firstDate.Day == AppConfig.IrregularStartDay;

        var extractedRows = _biometricService.ExtractEmployeeAttendance(
            _allBiometricRecords!, employee.EmployeeId, startIndex, _calendar.TotalDays, isIrregular);

        var dailyRecords = extractedRows.Select(ToDailyAttendanceRecord).ToList();

        var context = new AttendanceContext(employee, _calendar);
        var analyzer = new AttendanceAnalyzer();
        analyzer.Analyze(dailyRecords, context);

        _reportGenerator.WriteEmployeeAttendance(employee, context, _calendar);

        if (_saveExtractedFilesPerEmployee)
        {
            SaveExtractedAttendanceFile(employee, extractedRows);
        }

        _processedCount++;
    }

    private static DailyAttendanceRecord ToDailyAttendanceRecord(BiometricRecord raw) => new()
    {
        AttendanceDate = raw.Date,
        ClockIn = raw.ClockIn,
        ClockOut = raw.ClockOut,
        Status = raw.Status,
        TotalWorkDurationHours = raw.TotalWorkDurationHours
    };

    /// <summary>
    /// Saves each employee's extracted (already payroll-cycle-aligned) raw
    /// attendance rows to their own .xlsx file, mirroring
    /// SaveExtractedBiometricAttendance in the VBA version. This gives you an
    /// audit trail per employee, independent of the summary report.
    ///
    /// Columns written (in order):
    ///   1 Date | 2 Day | 3 Clock In | 4 Clock Out | 5 Status | 6 Total Work Duration
    ///
    /// Total Work Duration is written as hours (a plain number) rather than as
    /// a formatted "hh:mm:ss" cell, because that is how it is consumed by the
    /// analyzer (DailyAttendanceRecord.TotalWorkDurationHours is a double of
    /// hours). If you need the audit file to look like the source sheet, swap
    /// the value expression for a DateTime-based hh:mm:ss formatting instead.
    /// </summary>
    private void SaveExtractedAttendanceFile(Employee employee, List<BiometricRecord> extractedRows)
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.Worksheets.Add("Attendance");

        sheet.Cell(1, 1).Value = "Date";
        sheet.Cell(1, 2).Value = "Day";
        sheet.Cell(1, 3).Value = "Clock In";
        sheet.Cell(1, 4).Value = "Clock Out";
        sheet.Cell(1, 5).Value = "Status";
        sheet.Cell(1, 6).Value = "Total Work Duration (hrs)";
        sheet.Row(1).Style.Font.Bold = true;

        int row = 2;
        foreach (var record in extractedRows)
        {
            sheet.Cell(row, 1).Value = record.Date;
            sheet.Cell(row, 1).Style.DateFormat.Format = "dd-mmm-yyyy";
            sheet.Cell(row, 2).Value = record.Date.DayOfWeek.ToString();

            if (record.ClockIn.HasValue)
            {
                sheet.Cell(row, 3).Value = DateTime.Today.Add(record.ClockIn.Value);
                sheet.Cell(row, 3).Style.DateFormat.Format = "hh:mm:ss";
            }
            if (record.ClockOut.HasValue)
            {
                sheet.Cell(row, 4).Value = DateTime.Today.Add(record.ClockOut.Value);
                sheet.Cell(row, 4).Style.DateFormat.Format = "hh:mm:ss";
            }

            sheet.Cell(row, 5).Value = record.Status;

            if (record.TotalWorkDurationHours.HasValue)
            {
                sheet.Cell(row, 6).Value = record.TotalWorkDurationHours.Value;
                sheet.Cell(row, 6).Style.NumberFormat.Format = "0.0000";
            }

            row++;
        }

        sheet.Columns().AdjustToContents();
        workbook.SaveAs(_fileService.GenerateFileName(employee));
    }
}