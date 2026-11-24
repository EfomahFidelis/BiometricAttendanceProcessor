namespace BiometricAttendanceProcessor.Services;

/// <summary>
/// Encapsulates everything the rest of the program needs to know about "when is
/// the payroll cycle, and which dates in it are public holidays" - equivalent to
/// the VBA "CPayrollCalendar" class.
///
/// DESIGN NOTE: notice there is no IsWeekend(date) method here, unlike the VBA
/// version. That's deliberate - DateTime.DayOfWeek already tells you that
/// natively (see DailyAttendanceRecord.IsWeekend), so a wrapper method would just
/// be redundant surface area to maintain.
/// </summary>
public class PayrollCalendar
{
    public DateTime StartDate { get; }
    public DateTime EndDate { get; }

    // Two collections built from the same holiday list, for two different jobs:
    //   - _holidaySet: O(1) "is this date a holiday?" lookups (HashSet).
    //   - _orderedHolidays: preserves the ORIGINAL ARRAY ORDER from AppConfig, so
    //     we can answer "which position (1st, 2nd, 3rd...) is this holiday in?" -
    //     needed to populate the named-holiday report columns (Eid Day 1, Eid Day
    //     2, Good Friday, Easter Monday) in the same order Management lists them.
    private readonly HashSet<DateTime> _holidaySet;
    private readonly List<DateTime> _orderedHolidays;

    public int TotalDays => (EndDate.Date - StartDate.Date).Days + 1;

    public PayrollCalendar(DateTime startDate, DateTime endDate, IEnumerable<DateTime> publicHolidays)
    {
        StartDate = startDate.Date;
        EndDate = endDate.Date;
        _orderedHolidays = publicHolidays.Select(d => d.Date).ToList();
        _holidaySet = new HashSet<DateTime>(_orderedHolidays);
    }

    public bool IsPublicHoliday(DateTime date) => _holidaySet.Contains(date.Date);

    public bool IsDateInRange(DateTime date) => date.Date >= StartDate && date.Date <= EndDate;

    /// <summary>Every calendar date in the payroll cycle, in order, inclusive of both endpoints.</summary>
    public IEnumerable<DateTime> GetPayrollDates()
    {
        for (var d = StartDate; d <= EndDate; d = d.AddDays(1))
            yield return d;
    }

    /// <summary>
    /// Returns the 1-based position of the given date within the configured
    /// public-holiday list (1 = first holiday in AppConfig.PublicHolidays, etc.),
    /// or -1 if the date is not a configured holiday. Used by
    /// AttendanceReportGenerator to write counts into the correct named-holiday
    /// column, mirroring the "holidayCounter" offset loop in the original VBA
    /// Write_Public_Holiday_Details routine.
    /// </summary>
    public int GetPublicHolidayIndex(DateTime date)
    {
        var index = _orderedHolidays.IndexOf(date.Date);
        return index == -1 ? -1 : index + 1;
    }
}
