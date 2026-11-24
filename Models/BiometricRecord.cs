namespace BiometricAttendanceProcessor.Models;

/// <summary>
/// A raw row exactly as it appears in the "DataSheet" worksheet, before any
/// business-rule analysis has been applied. This is intentionally a plain,
/// "dumb" data holder (sometimes called a DTO - Data Transfer Object) with no
/// behaviour of its own.
///
/// WHY TWO SEPARATE MODELS (BiometricRecord vs DailyAttendanceRecord)?
/// This separation mirrors a distinction the VBA code blurred together:
///   - BiometricRecord = "what the spreadsheet says" (raw, unopinionated data).
///   - DailyAttendanceRecord = "what the business rules say about that data"
///     (IsPresent, IsWeekend, IsPublicHoliday, etc. - all derived/computed).
/// Keeping them separate means BiometricDataService (which only knows how to
/// read Excel) never needs to know anything about attendance business rules,
/// and AttendanceAnalyzer (which only knows business rules) never needs to know
/// anything about Excel. Each class has exactly one reason to change.
/// </summary>
public class BiometricRecord
{
    public string EmployeeId { get; set; } = string.Empty;
    public DateTime Date { get; set; }
    public TimeSpan? ClockIn { get; set; }
    public TimeSpan? ClockOut { get; set; }
    public string Status { get; set; } = string.Empty;

    /// <summary>
    /// Total Work Duration from Column J, in hours. This is the AUTHORITATIVE
    /// duration figure from the biometric export - it is correct even for
    /// overnight "Shift Change" shifts where ClockOut is numerically less than
    /// ClockIn. Null when the source cell was blank or "N.A.".
    /// </summary>
    public double? TotalWorkDurationHours { get; set; }

    /// <summary>The 1-based row number this came from in DataSheet - useful for diagnostics.</summary>
    public int SourceRowNumber { get; set; }
}
