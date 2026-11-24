namespace BiometricAttendanceProcessor.Configuration;

/// <summary>
/// Central, static configuration for a single payroll run.
///
/// This is the direct equivalent of the VBA "MConfigSettings" standard module.
/// The reasoning for keeping this separate from the rest of the code is the same
/// reasoning we used in the VBA version: the values in the "MODIFY THESE EVERY
/// PAYROLL CYCLE" block below are the ONLY things a non-programmer (or a junior
/// developer) should ever need to touch when a new month's payroll run begins.
/// Nothing else in the codebase should need to change from month to month.
///
/// NOTE ON DESIGN: a static class with public const/readonly fields is the
/// simplest possible way to express "global configuration" in C#. For a larger
/// production system you would typically load these values from appsettings.json
/// via the Microsoft.Extensions.Configuration package instead of hard-coding them
/// here, which would let you change the payroll dates without recompiling the
/// program at all. That's flagged as a natural next step in the README.
/// </summary>
public static class AppConfig
{
    // ============================================================
    // MODIFY THESE EVERY PAYROLL CYCLE
    // ============================================================

    /// <summary>First day of the payroll cycle (inclusive). E.g. 19-Mar-2026.</summary>
    public static readonly DateTime PayrollStartDate = new(2026, 9, 19);

    /// <summary>Last day of the payroll cycle (inclusive). E.g. 18-Apr-2026.</summary>
    public static readonly DateTime PayrollEndDate = new(2026, 10, 18);

    /// <summary>
    /// Public holidays that fall within the current payroll cycle. Add or remove
    /// dates here as the cycle changes - nothing else in the codebase references
    /// specific holiday dates directly. Comments record which named holiday each
    /// date represents, purely for human readability; the program itself only
    /// cares about the date, not the name (see PayrollCalendar.GetPublicHolidayIndex
    /// for how the *position* in this array maps to the named report columns).
    /// </summary>
    public static readonly DateTime[] PublicHolidays =
    {
        new DateTime(2026, 10, 01), // Eid el Fitri Day 1
        // new DateTime(2026, 3, 20), // Eid-El-Fitri Day 2
        // new DateTime(2026, 4, 3),  // Good Friday
        // new DateTime(2026, 4, 6),  // Easter Monday
    };

    // ============================================================
    // BUSINESS RULE THRESHOLDS - change only when Management updates policy
    // ============================================================

    /// <summary>
    /// Minimum hours of work required, on a public holiday or a weekend, before
    /// the day is paid in full ("PhP"/"P") rather than partially ("3HL"). Current
    /// policy (per Management, both Consultant and Contractor): at least 3 hours,
    /// matching the VBA IsPublicHolidayWorkDurationEligible rule ("&gt;= means
    /// that exactly the required number of hours qualifies"). See
    /// DailyAttendanceRecord.IsWorkDurationAtLeast for the comparison this drives.
    /// </summary>
    public const double MinimumOtHours = 3.0;

    // ============================================================
    // STRUCTURAL CONSTANTS - should rarely, if ever, need to change
    // ============================================================

    /// <summary>
    /// If an employee's biometric data starts on this day-of-month, it means one
    /// extra stray day from the previous payroll cycle crept into the export and
    /// must be discarded before extraction begins. See BiometricDataService for
    /// the extraction algorithm this drives.
    /// </summary>
    public const int IrregularStartDay = 18;

    /// <summary>The day-of-month a payroll cycle is supposed to start on.</summary>
    public const int RegularStartDay = 19;

    public const string SourceCriteriaSheetName = "CriteriaSheet";
    public const string SourceDataSheetName = "DataSheet";
    public const string ReportSheetName = "Extraction Records";
}