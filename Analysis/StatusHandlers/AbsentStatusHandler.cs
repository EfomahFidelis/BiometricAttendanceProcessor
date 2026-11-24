using BiometricAttendanceProcessor.Models;

namespace BiometricAttendanceProcessor.Analysis.StatusHandlers;

/// <summary>
/// Handles the "Absent" status token, per the specialized absenteeism rules:
///
/// Contractor:
///   - Absent on a non-holiday day -> penalized, coded "X".
///   - Absent on a public holiday -> never penalized, coded "XPH".
///   - Absent + a time-off token also present (should never happen in practice,
///     but defended against) -> Absent wins, never penalized, coded "XTO".
///
/// Consultant:
///   - Absent Mon-Fri, non-holiday -> penalized, coded "X".
///   - Absent on a public holiday -> never penalized, coded "XPH".
///   - Absent on a weekend -> never penalized (it's their legitimate weekly-off),
///     coded "XWO". Note this is a DIFFERENT code from the plain "W/O" produced
///     by TimeOffStatusHandler: "XWO" specifically means "flagged absent, but
///     exempted because it's the weekend", whereas "W/O" means "this day was
///     never in question - it was a scheduled day off from the start."
/// </summary>
public class AbsentStatusHandler : IAttendanceStatusHandler
{
    public void Handle(DailyAttendanceRecord record, AttendanceContext context)
    {
        var employee = context.Employee;

        // The public-holiday exemption applies identically to both contract
        // types, so it's checked once, first, before any contract-type branch -
        // this guarantees it can never be bypassed by either branch below.
        if (record.IsPublicHoliday)
        {
            context.StoreCode(record.AttendanceDate, "XPH");
            return;
        }

        switch (employee.Type)
        {
            case Models.ContractType.Contractor:
                HandleContractor(record, context);
                break;

            case Models.ContractType.Consultant:
                HandleConsultant(record, context);
                break;
        }
    }

    private static void HandleContractor(DailyAttendanceRecord record, AttendanceContext context)
    {
        if (ContainsTimeOffToken(record.Status))
        {
            // Defensive case: "Absent" and a time-off token should never
            // co-occur per business rules. If it somehow does, Absent takes
            // precedence in the code, but the employee is NOT penalized.
            context.StoreCode(record.AttendanceDate, "XTO");
        }
        else
        {
            context.StoreCode(record.AttendanceDate, "X");
            context.Employee.AbsentCount++;
        }
    }

    private static void HandleConsultant(DailyAttendanceRecord record, AttendanceContext context)
    {
        if (record.IsWeekend)
        {
            context.StoreCode(record.AttendanceDate, "XWO");
        }
        else
        {
            context.StoreCode(record.AttendanceDate, "X");
            context.Employee.AbsentCount++;
        }
    }

    private static bool ContainsTimeOffToken(string status)
    {
        var tokens = AttendanceStatusParser.Parse(status);
        return tokens.Contains("Weekly Off") || tokens.Contains("WeeklyOff") || tokens.Contains("On Time Off");
    }
}
