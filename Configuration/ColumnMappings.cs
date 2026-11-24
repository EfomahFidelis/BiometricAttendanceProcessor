namespace BiometricAttendanceProcessor.Configuration;

/// <summary>
/// Column layout of the "CriteriaSheet" worksheet (one row per employee).
/// Values are 1-based to match ClosedXML's Cell(row, column) indexing, which in
/// turn matches Excel's own column numbering (A=1, B=2, ...) - so these numbers
/// read the same way the VBA "CriteriaColumns" Enum did.
/// </summary>
public enum CriteriaColumn
{
    EmployeeId = 1,   // Column A
    EmployeeName = 2, // Column B
    Location = 3,     // Column C
    JobTitle = 4,     // Column D
    ContractType = 5, // Column E
    StartDate = 6,    // Column F
    EndDate = 7       // Column G
}

/// <summary>
/// Column layout of the "DataSheet" worksheet (raw biometric export).
///
/// These are ABSOLUTE Excel column positions (A=1, B=2, ...), not offsets into
/// the 6-column extract the VBA code builds from Range("E2:J..."). The two
/// numbering schemes describe the same physical cells:
///
///   VBA extract index (1..6)   This enum              Physical column
///   ------------------------   ----------------------  -----------------
///   1 (Date)                   Date = 5                E
///   2 (Day)                    Day = 6                 F
///   3 (Clock In)               ClockIn = 7             G
///   4 (Clock Out)              ClockOut = 8            H
///   5 (Status)                 Status = 9              I
///   6 (Total Work Duration)    TotalWorkDuration = 10  J
///
/// The old enum had Date/Day/ClockIn/ClockOut/Status shifted one column to the
/// right (it counted a phantom "Shift" column at D that isn't in the real
/// sheet, and it had no notion of Total Work Duration at all - it called
/// Column J "Remarks"). This version matches the real header row:
///   A Employee Id | B Employee Name | C Employee Sub Type | D Current Department
///   E Date | F Day | G Clock In | H Clock Out | I Status | J Total Work Duration
///   K Detailed Status | L Office Location | M Paying Org | N Accounting Org
/// </summary>
public enum BiometricColumn
{
    EmployeeId = 1,           // Column A
    Name = 2,                 // Column B
    EmployeeSubType = 3,      // Column C
    CurrentDepartment = 4,    // Column D
    Date = 5,                 // Column E
    Day = 6,                  // Column F
    ClockIn = 7,              // Column G
    ClockOut = 8,             // Column H
    Status = 9,               // Column I
    TotalWorkDuration = 10,   // Column J  <-- the authoritative duration column
    DetailedStatus = 11,      // Column K
    // Columns L, M, N (Office Location, Paying Org, Accounting Org) exist in the
    // source but are not consumed by this program. Add them here if that changes.
}

/// <summary>
/// Column layout of the generated "Extraction Records" report worksheet.
/// AttendanceGridStart marks the first of the day-by-day attendance code columns
/// (column AE); each subsequent day is written one column to the right of it.
/// </summary>
public enum ReportColumn
{
    EmployeeId = 1,
    EmployeeName = 2,
    ContractType = 3,
    Mandays = 4,
    Present = 5,
    MissingDays = 6,
    ClockIns = 7,
    ClockOuts = 8,
    ClockInOuts = 9,
    Absent = 10,
    PresentOnLeave = 11,
    Leave = 12,
    WeeklyOff = 13,
    SunPresent = 14,
    MonPresent = 15,
    TuePresent = 16,
    WedPresent = 17,
    ThuPresent = 18,
    FriPresent = 19,
    SatPresent = 20,
    OtSunday = 21,
    OtSaturday = 22,
    PublicHolidayCount = 23,
    // Columns 24-30 (X through AD) hold the named-holiday breakdown and are
    // written as PublicHolidayCount + index (1..4) - see AttendanceReportGenerator.
    AttendanceGridStart = 31 // Column AE - first day-of-month attendance code cell
}