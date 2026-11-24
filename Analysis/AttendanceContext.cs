using BiometricAttendanceProcessor.Models;
using BiometricAttendanceProcessor.Services;

namespace BiometricAttendanceProcessor.Analysis;

/// <summary>
/// Holds all the mutable, per-employee working state that accumulates while a
/// month's worth of daily records is walked, one day at a time.
///
/// DESIGN NOTE: In the VBA "Class Approach", this state was split across three
/// private fields on CAttendanceAnalyzer (m_AttendanceRecords, a Dictionary
/// keyed by formatted date strings; m_PublicHolidayCounts, another Dictionary;
/// and m_DayPresentCounts, a 1-to-7 array) plus the separately-passed Employee
/// object. Every status handler method reached into whichever of those it
/// needed directly, via module-level scope.
///
/// Here, all of that state is gathered into one explicit object that gets
/// passed as a parameter to every IAttendanceStatusHandler.Handle() call. This
/// makes the DATA DEPENDENCY VISIBLE at every call site - you can look at a
/// handler's method signature and immediately know everything it's allowed to
/// read or write, rather than having to trust that "some global variable
/// somewhere" hasn't been touched by a different code path. This directly
/// addresses the class of bugs we spent time on in VBA, where one handler
/// silently double-counted a value another handler had already set.
/// </summary>
public class AttendanceContext
{
    public Employee Employee { get; }
    public PayrollCalendar Calendar { get; }

    /// <summary>Date -> accumulated attendance code string (e.g. "P", "PhP", "XWO").</summary>
    public Dictionary<DateTime, string> AttendanceCodes { get; } = new();

    /// <summary>Date -> number of confirmed, paid public-holiday attendances on that date.</summary>
    public Dictionary<DateTime, int> PublicHolidayCounts { get; } = new();

    /// <summary>
    /// Index 0 = Sunday ... 6 = Saturday, matching the numeric value of C#'s
    /// built-in DayOfWeek enum exactly, so no translation step is ever needed
    /// between "which day is this" and "which array slot do I increment".
    /// </summary>
    public int[] DayPresentCounts { get; } = new int[7];

    public AttendanceContext(Employee employee, PayrollCalendar calendar)
    {
        Employee = employee;
        Calendar = calendar;
    }

    /// <summary>
    /// Appends a code to whatever is already stored for this date (matching the
    /// VBA UpdateAttendance behaviour of concatenating, e.g. "3HL" + "3HL" if
    /// called twice - which is exactly the bug we fixed by ensuring only ONE
    /// status handler ever runs per day; see AttendanceAnalyzer.DispatchStatus).
    /// </summary>
    public void StoreCode(DateTime date, string code)
    {
        if (AttendanceCodes.TryGetValue(date, out var existing))
            AttendanceCodes[date] = existing + code;
        else
            AttendanceCodes[date] = code;
    }

    public string GetCode(DateTime date) =>
        AttendanceCodes.TryGetValue(date, out var code) ? code : string.Empty;

    public void IncrementPublicHolidayCount(DateTime date)
    {
        PublicHolidayCounts[date] = PublicHolidayCounts.TryGetValue(date, out var count) ? count + 1 : 1;
    }

    public void IncrementDayPresentCount(DayOfWeek day)
    {
        DayPresentCounts[(int)day]++;
    }
}
