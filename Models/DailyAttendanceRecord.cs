namespace BiometricAttendanceProcessor.Models;

/// <summary>
/// Represents a single day's biometric attendance row, plus every boolean flag
/// derived from it, for exactly one employee.
///
/// This mirrors the VBA "CDailyAttendanceRecord" class - but notice how much
/// smaller and safer it is here. Two specific VBA bugs we tracked down are
/// structurally impossible in this version:
///
///   1. The "IsDate() on a Variant/Double always returns False" bug: VBA stored
///      clock times as ambiguous Variants and had to guess, at read time,
///      whether the underlying value could be coerced to a Date. Here, ClockIn
///      and ClockOut are declared as `TimeSpan?` (nullable TimeSpan) - there is
///      no ambiguity. A value either IS a valid time (HasValue == true) or it
///      ISN'T (HasValue == false, meaning blank / "N.A." in the source data).
///
///   2. The "day-of-week string matching" bugs (Direct Approach's gridVisual
///      priority-order issues, and comparisons like `dayOfWeek = "Saturday"`):
///      C#'s built-in DayOfWeek enum (Sunday=0 ... Saturday=6) comes directly off
///      DateTime.DayOfWeek, computed correctly by the runtime from the date
///      itself. There's no possibility of the "Day" column text disagreeing with
///      the actual calendar date, because we never trust the raw "Day" text at
///      all - we derive it from AttendanceDate.
/// </summary>
public class DailyAttendanceRecord
{
    public DateTime AttendanceDate { get; set; }

    /// <summary>Nullable clock-in time. Null means "no valid clock-in on this row".</summary>
    public TimeSpan? ClockIn { get; set; }

    /// <summary>Nullable clock-out time. Null means "no valid clock-out on this row".</summary>
    public TimeSpan? ClockOut { get; set; }

    /// <summary>The raw, comma-separated status text exactly as read from the source, e.g. "Present,Holiday,Request".</summary>
    public string Status { get; set; } = string.Empty;

    public bool IsValidDate { get; set; }

    /// <summary>
    /// Set by the AttendanceAnalyzer (via PayrollCalendar.IsPublicHoliday) before
    /// any status handler runs, since several handlers need to know this.
    /// </summary>
    public bool IsPublicHoliday { get; set; }

    /// <summary>
    /// Total Work Duration from Column J of the source DataSheet, in hours.
    /// This is the AUTHORITATIVE duration figure from the biometric export - it
    /// is correct even for overnight "Shift Change" shifts where ClockOut is
    /// numerically less than ClockIn. Null when the source cell was blank or
    /// "N.A.". Populated by BiometricDataService via AttendanceProcessor.
    /// </summary>
    public double? TotalWorkDurationHours { get; set; }

    /// <summary>Derived directly from the date - see class remarks above.</summary>
    public DayOfWeek DayOfWeek => AttendanceDate.DayOfWeek;

    public bool IsWeekend => DayOfWeek is DayOfWeek.Saturday or DayOfWeek.Sunday;

    public bool IsClockIn => ClockIn.HasValue;
    public bool IsClockOut => ClockOut.HasValue;

    /// <summary>Present means a genuine clock-in AND a genuine clock-out were both recorded.</summary>
    public bool IsPresent => IsClockIn && IsClockOut;

    /// <summary>Clocked in but never clocked out - a "Single Punch Absent" candidate.</summary>
    public bool IsSinglePunchAbsent => IsClockIn && !IsClockOut;

    /// <summary>
    /// Hours actually worked. Prefers the authoritative Column J value (which is
    /// correct for overnight "Shift Change" shifts). Falls back to
    /// ClockOut - ClockIn only when Column J is missing, and wraps a negative
    /// result by 24 hours so an overnight fallback still reports a sensible
    /// positive duration. Returns 0 if neither source is available.
    ///
    /// The fallback exists purely as a defensive measure for malformed rows -
    /// in normal operation Column J is always present.
    /// </summary>
    public double GetWorkDurationHours()
    {
        if (TotalWorkDurationHours.HasValue)
            return TotalWorkDurationHours.Value;

        if (ClockIn.HasValue && ClockOut.HasValue)
        {
            var diff = (ClockOut.Value - ClockIn.Value).TotalHours;
            if (diff < 0) diff += 24; // overnight wrap, defensive fallback only
            return diff;
        }
        return 0;
    }

    /// <summary>
    /// "At least" comparison, matching the VBA IsPublicHolidayWorkDurationEligible
    /// rule: "&gt;= means that exactly the required number of hours qualifies."
    /// Driven by AppConfig.MinimumOtHours (currently 3.0 hours).
    /// </summary>
    public bool IsWorkDurationAtLeast(double minHours) => GetWorkDurationHours() >= minHours;
}