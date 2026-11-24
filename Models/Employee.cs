namespace BiometricAttendanceProcessor.Models;

/// <summary>
/// A closed set of the two contract types the business rules distinguish between.
///
/// DESIGN NOTE: The VBA code compared raw strings everywhere - "Consultant",
/// "Contractor" - which is fragile (a stray space or a typo like "Consultants"
/// silently falls through every "Case" branch with no compiler warning). An enum
/// makes this a compile-time-checked concept: ContractType.Cosnultant simply
/// won't compile, whereas "Cosnultant" as a string would compile and fail silently
/// at runtime. This is one of the clearest wins of moving from VBA to C#.
/// </summary>
public enum ContractType
{
    Unknown,
    Consultant,
    Contractor
}

/// <summary>Converts the raw text found in the CriteriaSheet into a ContractType.</summary>
public static class ContractTypeParser
{
    public static ContractType Parse(string? raw) => raw?.Trim() switch
    {
        "Consultant" => ContractType.Consultant,
        "Contractor" => ContractType.Contractor,
        _ => ContractType.Unknown
    };
}

/// <summary>
/// Represents one employee's identity plus every running total accumulated while
/// their attendance is analyzed for the current payroll cycle.
///
/// This is the equivalent of the VBA "dataArrayCurrentEmployeeDetail" Variant
/// array (an 18-element array addressed by "magic numbers" like
/// dataArrayCurrentEmployeeDetail(11) for AbsentCount). Using a class with named
/// properties instead of a numerically-indexed array eliminates an entire
/// category of bugs we spent time chasing in the VBA version - e.g. writing to
/// the wrong index, or two different pieces of code disagreeing about what
/// index 14 means.
/// </summary>
public class Employee
{
    // --- Identity (read from CriteriaSheet, set once) ---
    public string EmployeeId { get; set; } = string.Empty;
    public string Name { get; set; } = string.Empty;
    public string Location { get; set; } = string.Empty;
    public string JobTitle { get; set; } = string.Empty;
    public ContractType Type { get; set; }
    public long WorkingDaysInMonth { get; set; }

    // --- Clock-machine activity totals (updated once per attendance row) ---
    public int PresentCount { get; set; }
    public int ClockInCount { get; set; }
    public int ClockOutCount { get; set; }
    public int ClockInOutCount { get; set; }

    // --- Attendance status totals ---
    public int AbsentCount { get; set; }
    public int SinglePunchAbsentCount { get; set; }

    /// <summary>Genuine paid/unpaid leave days (fixed-sum payment, e.g. 7 or 14 days).</summary>
    public int LeaveCount { get; set; }

    /// <summary>
    /// Days the employee was scheduled for leave but actually worked - i.e. they
    /// forfeited the leave day and are entitled to per-day encashment instead of
    /// the fixed leave sum. Stored separately from LeaveCount; see
    /// StatusHandlers/LeaveStatusHandler for where this is populated.
    /// </summary>
    public int WorkingOnLeaveCount { get; set; }

    public int WeeklyOffCount { get; set; }
    public int OtSundayCount { get; set; }
    public int OtSaturdayCount { get; set; }
    public int MissingBlankDaysCount { get; set; }

    /// <summary>
    /// True when the employee has a genuine, non-placeholder ID. Equivalent of
    /// VBA's "dataArrayCurrentEmployeeDetail(1) <> No EmpID" checks, but exposed
    /// here as a single, unambiguous, testable property instead of a string
    /// comparison repeated at every call site.
    /// </summary>
    public bool HasValidId => !string.IsNullOrWhiteSpace(EmployeeId) && EmployeeId != "No EmpID";
}
