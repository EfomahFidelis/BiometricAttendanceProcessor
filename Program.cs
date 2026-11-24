using ClosedXML.Excel;
using BiometricAttendanceProcessor.Configuration;
using BiometricAttendanceProcessor.Processing;
using BiometricAttendanceProcessor.Reporting;
using BiometricAttendanceProcessor.Services;

namespace BiometricAttendanceProcessor;

/// <summary>
/// THE COMPOSITION ROOT.
///
/// This is the ONE place in the entire program where every concrete class gets
/// instantiated and wired together. Every other class in the project depends
/// only on interfaces (IEmployeeRepository, IBiometricDataService) or on other
/// plain classes passed in via their constructors - none of them reach out and
/// create their own dependencies. That discipline (called "Dependency
/// Injection" when done via constructors, as it is here) is what makes it
/// possible to swap any one piece - say, replacing EmployeeRepository with a
/// database-backed version - by changing only the two or three lines below,
/// with zero changes anywhere else in the codebase.
/// </summary>
public static class Program
{
    public static void Main(string[] args)
    {
        Console.WriteLine("=== Biometric Attendance Processor ===");
        Console.WriteLine();

        string sourcePath = args.Length > 0 ? args[0] : PromptForInput("Path to source workbook (.xlsx): ");
        string outputFolder = args.Length > 1 ? args[1] : PromptForInput("Output folder for reports: ");

        if (!File.Exists(sourcePath))
        {
            Console.WriteLine($"ERROR: source file not found: {sourcePath}");
            return;
        }

        Directory.CreateDirectory(outputFolder);

        using var workbook = new XLWorkbook(sourcePath);

        var criteriaSheet = workbook.Worksheet(AppConfig.SourceCriteriaSheetName);
        var dataSheet = workbook.Worksheet(AppConfig.SourceDataSheetName);

        // Re-use an existing "Extraction Records" sheet if present (clearing old
        // data first), otherwise create a fresh one - matching the VBA
        // Init_Report behaviour of clearing A4:BZ400 before each run.
        IXLWorksheet reportSheet = workbook.Worksheets.Contains(AppConfig.ReportSheetName)
            ? workbook.Worksheet(AppConfig.ReportSheetName)
            : workbook.Worksheets.Add(AppConfig.ReportSheetName);

        reportSheet.Range("A4:BZ400").Clear();

        // ---- Composition: build each collaborator, then hand them to the Facade ----
        var calendar = new PayrollCalendar(AppConfig.PayrollStartDate, AppConfig.PayrollEndDate, AppConfig.PublicHolidays);
        var employeeRepository = new EmployeeRepository(criteriaSheet);
        var biometricService = new BiometricDataService(dataSheet);
        var reportGenerator = new AttendanceReportGenerator(reportSheet);
        var fileService = new FileSystemService(outputFolder);

        var processor = new AttendanceProcessor(
            employeeRepository, biometricService, reportGenerator, fileService, calendar);

        processor.ProcessAll();

        // EXTENSION POINT: the VBA version prompted "Do you want the range
        // sorted?" and applied a multi-column sort to the report before saving.
        // That's intentionally left out here - ClosedXML's IXLRange.Sort API
        // can replicate it (SortFields.Add per column, similar to the VBA
        // SortWorksheet routine) if you want to add it back in; this is a good
        // first modification to try on your own.

        var reportPath = Path.Combine(outputFolder, "AttendanceReport.xlsx");
        workbook.SaveAs(reportPath);

        Console.WriteLine();
        Console.WriteLine($"Report saved to: {reportPath}");
        Console.WriteLine($"Per-employee extracted attendance files saved to: {fileService.OutputFolder}");
    }

    private static string PromptForInput(string prompt)
    {
        Console.Write(prompt);
        var input = Console.ReadLine() ?? string.Empty;
        return input.Trim().Trim('"');
    }
}
