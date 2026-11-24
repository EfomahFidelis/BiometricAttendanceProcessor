using BiometricAttendanceProcessor.Configuration;
using BiometricAttendanceProcessor.Models;

namespace BiometricAttendanceProcessor.Analysis.StatusHandlers;

/// <summary>
/// These four handlers are grouped in one file purely because each is short and
/// self-contained - feel free to split them into individual files as the
/// project grows, which is generally the more common C# convention (one public
/// type per file).
/// </summary>

/// <summary>Handles the literal "Unpaid Leave" status token.</summary>
public class UnpaidLeaveStatusHandler : IAttendanceStatusHandler
{
    public void Handle(DailyAttendanceRecord record, AttendanceContext context)
    {
        context.Employee.LeaveCount++;
        context.StoreCode(record.AttendanceDate, "uL");
    }
}

/// <summary>
/// Handles a clock-in with no matching clock-out. Always counts against the
/// absence total, regardless of contract type - there is no Direct-Approach
/// equivalent that exempts this case.
/// </summary>
public class SinglePunchAbsentStatusHandler : IAttendanceStatusHandler
{
    public void Handle(DailyAttendanceRecord record, AttendanceContext context)
    {
        context.StoreCode(record.AttendanceDate, "SPA");
        context.Employee.AbsentCount++;
    }
}

/// <summary>
/// Handles the "Leave" status token.
///
/// KEY DISTINCTION (this is the fix for the "Present on Leave" bug):
///   - Employee actually worked that day despite being scheduled for leave ->
///     they forfeited the leave day and are entitled to per-day encashment.
///     Coded "PL", counted in Employee.WorkingOnLeaveCount.
///   - Employee did not work -> genuine leave, paid as a fixed lump sum per
///     their contract rotation. Coded "L", counted in Employee.LeaveCount.
/// These are two DIFFERENT counters feeding two DIFFERENT report columns -
/// mixing them up (as the original buggy version did, always incrementing
/// LeaveCount regardless of which code was stored) misreports how much the
/// employee is owed.
/// </summary>
public class LeaveStatusHandler : IAttendanceStatusHandler
{
    public void Handle(DailyAttendanceRecord record, AttendanceContext context)
    {
        var employee = context.Employee;

        if (record.IsPresent)
        {
            context.StoreCode(record.AttendanceDate, "PL");
            employee.WorkingOnLeaveCount++;
        }
        else
        {
            context.StoreCode(record.AttendanceDate, "L");
            employee.LeaveCount++;
        }
    }
}

/// <summary>
/// Handles both "Weekly Off"/"WeeklyOff" and "On Time Off" status tokens. The
/// two tokens differ only in their DISPLAY code ("W/O" vs "T/O") - the
/// underlying business rules are identical, so one handler covers both,
/// determining which display code to use by inspecting the raw status text.
/// This replaces two separate VBA call sites (ProcessAttendanceStatus_WeeklyOff
/// and the "On Time Off" case) that both funnelled into the same
/// ProcessAttendanceStatus_TimeOff routine anyway.
/// </summary>
public class TimeOffStatusHandler : IAttendanceStatusHandler
{
    public void Handle(DailyAttendanceRecord record, AttendanceContext context)
    {
        var tokens = AttendanceStatusParser.Parse(record.Status);
        string displayCode = tokens.Contains("On Time Off") ? "T/O" : "W/O";
        var employee = context.Employee;

        switch (employee.Type)
        {
            case ContractType.Contractor:
                HandleContractor(record, context, displayCode);
                break;

            case ContractType.Consultant:
                HandleConsultant(record, context, displayCode);
                break;

            default:
                // Fallback for any employee whose contract type wasn't
                // recognized - matches the VBA "Case Else" default behaviour.
                context.StoreCode(record.AttendanceDate, displayCode);
                employee.WeeklyOffCount++;
                break;
        }
    }

    private static void HandleContractor(DailyAttendanceRecord record, AttendanceContext context, string displayCode)
    {
        if (record.IsPresent)
        {
            context.StoreCode(record.AttendanceDate, "P");
        }
        else
        {
            context.StoreCode(record.AttendanceDate, displayCode);
            context.Employee.WeeklyOffCount++;
        }
    }

    private static void HandleConsultant(DailyAttendanceRecord record, AttendanceContext context, string displayCode)
    {
        if (!record.IsWeekend)
        {
            // A Consultant tagged "Weekly Off"/"On Time Off" on a WEEKDAY is
            // still treated as a day off, regardless of presence.
            context.StoreCode(record.AttendanceDate, displayCode);
            context.Employee.WeeklyOffCount++;
            return;
        }

        // On an actual weekend, a Consultant's presence is evaluated as
        // potential overtime rather than a plain day-off.
        if (!record.IsPresent)
        {
            context.StoreCode(record.AttendanceDate, displayCode);
            context.Employee.WeeklyOffCount++;
        }
        else if (record.IsWorkDurationAtLeast(AppConfig.MinimumOtHours))
        {
            context.StoreCode(record.AttendanceDate, "P");
        }
        else
        {
            context.StoreCode(record.AttendanceDate, "3HL");
        }
    }
}
