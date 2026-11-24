using BiometricAttendanceProcessor.Models;

namespace BiometricAttendanceProcessor.Services;

/// <summary>
/// Owns all file-system concerns: creating the timestamped output folder for a
/// run, and generating a safe, unique file name per employee. Equivalent of the
/// VBA "CFileSystemService" class.
///
/// Keeping this logic in one small class (rather than scattered through the
/// processor) means if you ever need to change the naming convention, or write
/// to a network path, or add a "dry run, don't actually write files" mode, there
/// is exactly one place to make that change.
/// </summary>
public class FileSystemService
{
    private readonly string _outputFolder;
    private int _noIdCounter;

    public FileSystemService(string baseFolder)
    {
        var timestamp = DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss");
        _outputFolder = Path.Combine(baseFolder, timestamp);
        Directory.CreateDirectory(_outputFolder);
    }

    public string OutputFolder => _outputFolder;

    /// <summary>Builds a safe, unique .xlsx file path for one employee's extracted attendance.</summary>
    public string GenerateFileName(Employee employee)
    {
        var safeId = CleanFileName(employee.EmployeeId);
        var safeName = CleanFileName(employee.Name);

        if (safeName.Length > 50) safeName = safeName[..50];

        string baseName;
        if (employee.EmployeeId == "No EmpID")
        {
            _noIdCounter++;
            baseName = string.IsNullOrWhiteSpace(safeName)
                ? $"No_EmpID_{_noIdCounter}"
                : $"No_EmpID_{_noIdCounter} - {safeName}";
        }
        else
        {
            baseName = $"{safeId} - {safeName}";
        }

        if (string.IsNullOrWhiteSpace(baseName.Trim('-', ' ')))
        {
            baseName = $"Employee_{Guid.NewGuid():N}";
        }

        return Path.Combine(_outputFolder, $"{baseName}.xlsx");
    }

    private static string CleanFileName(string input)
    {
        foreach (var invalidChar in Path.GetInvalidFileNameChars())
        {
            input = input.Replace(invalidChar, '_');
        }
        return System.Text.RegularExpressions.Regex.Replace(input, @"\s+", " ").Trim();
    }
}
