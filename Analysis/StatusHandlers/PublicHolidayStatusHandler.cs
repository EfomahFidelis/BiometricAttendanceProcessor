using BiometricAttendanceProcessor.Configuration;
using BiometricAttendanceProcessor.Models;

namespace BiometricAttendanceProcessor.Analysis.StatusHandlers;

/// <summary>
/// Handles the "Holiday" and "Public Holiday" status tokens for BOTH contract
/// types (the underlying policy - "must work more than 3 hours to be paid in
/// full" - applies identically to Consultants and Contractors here; only the
/// weekend-OT rule differs between them, and that lives in
/// AttendanceAnalyzer.ProcessWeekendOvertime instead).
/// </summary>
public class PublicHolidayStatusHandler : IAttendanceStatusHandler
{
    public void Handle(DailyAttendanceRecord record, AttendanceContext context)
    {
        // Defensive guard: if the calendar doesn't actually confirm this date as
        // a public holiday (e.g. the status text says "Holiday" but the date
        // isn't in AppConfig.PublicHolidays), flag it rather than silently
        // trusting the raw status text.
        if (!record.IsValidDate || !record.IsPublicHoliday)
        {
            context.StoreCode(record.AttendanceDate, "eXP");
            return;
        }

        // No Single-Punch-Absent penalty is allowed on a public holiday - the
        // employee isn't obligated to work it at all, so a lone clock-in with no
        // clock-out isn't held against them the way it would be on a normal day.
        if (record.IsSinglePunchAbsent)
        {
            context.StoreCode(record.AttendanceDate, "pSPA");
            context.Employee.SinglePunchAbsentCount++;
            return;
        }

        if (record.IsPresent)
        {
            if (record.IsWorkDurationAtLeast(AppConfig.MinimumOtHours))
            {
                context.StoreCode(record.AttendanceDate, "PhP"); // Paid holiday Present
                context.IncrementPublicHolidayCount(record.AttendanceDate);
            }
            else
            {
                // Worked, but under the minimum-hours threshold - partial credit only.
                context.StoreCode(record.AttendanceDate, "3HL");
            }
        }
        else
        {
            // Did not work the holiday at all - no penalty, simply marked absent-on-holiday.
            context.StoreCode(record.AttendanceDate, "XPH");
        }
    }
}
