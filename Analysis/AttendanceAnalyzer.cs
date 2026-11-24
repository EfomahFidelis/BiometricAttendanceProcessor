using BiometricAttendanceProcessor.Analysis.StatusHandlers;
using BiometricAttendanceProcessor.Configuration;
using BiometricAttendanceProcessor.Models;

namespace BiometricAttendanceProcessor.Analysis;

/// <summary>
/// Walks one employee's month of daily attendance records and populates an
/// AttendanceContext with every derived statistic and attendance code.
///
/// This class plays the role of both the VBA "CAttendanceAnalyzer" class AND
/// the composition root for the Strategy objects it uses (see constructor).
/// </summary>
public class AttendanceAnalyzer
{
    /// <summary>
    /// PRIORITY ORDER MATTERS. When a single day's raw status text contains
    /// multiple tokens (e.g. "Present,Holiday,Request"), only ONE token may
    /// "win" and determine that day's code - otherwise two handlers could both
    /// fire and double-stamp the day (this is exactly the "3HL3HL" duplicate-code
    /// bug we tracked down earlier). Earlier entries in this array outrank later
    /// ones. This ordering was reverse-engineered from, and now matches, the
    /// corrected "gridVisual" priority array on the Direct (working) VBA
    /// approach - "Holiday" specifically must outrank "Absent" and "Leave", or a
    /// public holiday tagged with a stray "Absent"/"Leave" token gets routed to
    /// the wrong handler entirely.
    /// </summary>
    private static readonly string[] PriorityLabels =
    {
        "Holiday",
        "Absent",
        "Unpaid Leave",
        "Single Punch Absent",
        "Leave",
        "Public Holiday",
        "On Time Off",
        "Weekly Off",
        "WeeklyOff",
        "Present"
    };

    // DESIGN PATTERN: STRATEGY (see IAttendanceStatusHandler.cs for the full
    // explanation). This dictionary is the "strategy lookup table" - each
    // status label maps to the ONE handler object responsible for it. Several
    // labels intentionally map to the SAME handler instance ("Holiday" and
    // "Public Holiday" both map to _publicHoliday; "Weekly Off"/"WeeklyOff"/
    // "On Time Off" all map to _timeOff) because they share identical business
    // rules, differing only in cosmetic display text.
    private readonly Dictionary<string, IAttendanceStatusHandler> _handlers;

    public AttendanceAnalyzer()
    {
        // MANUAL COMPOSITION ROOT: this constructor is where all the Strategy
        // objects get created and wired together. For a small object graph like
        // this, doing it by hand (rather than pulling in a full Dependency
        // Injection container such as Microsoft.Extensions.DependencyInjection)
        // keeps things simple and easy to trace. If this graph grows much
        // larger, introducing a DI container becomes worthwhile - flagged as a
        // natural next step in the README.
        var publicHoliday = new PublicHolidayStatusHandler();
        var absent = new AbsentStatusHandler();
        var present = new PresentStatusHandler(publicHoliday, absent);
        var timeOff = new TimeOffStatusHandler();
        var leave = new LeaveStatusHandler();
        var unpaidLeave = new UnpaidLeaveStatusHandler();
        var singlePunchAbsent = new SinglePunchAbsentStatusHandler();

        _handlers = new Dictionary<string, IAttendanceStatusHandler>
        {
            ["Holiday"] = publicHoliday,
            ["Public Holiday"] = publicHoliday,
            ["Absent"] = absent,
            ["Unpaid Leave"] = unpaidLeave,
            ["Single Punch Absent"] = singlePunchAbsent,
            ["Leave"] = leave,
            ["On Time Off"] = timeOff,
            ["Weekly Off"] = timeOff,
            ["WeeklyOff"] = timeOff,
            ["Present"] = present,
        };
    }

    /// <summary>
    /// Processes every daily record for one employee, mutating the supplied
    /// context (and, via context.Employee, the employee's running totals) as it
    /// goes. Nothing is returned - the context IS the accumulated result.
    /// </summary>
    public void Analyze(IEnumerable<DailyAttendanceRecord> records, AttendanceContext context)
    {
        foreach (var record in records)
        {
            // Set once per record, before any handler runs, since several
            // handlers need to know this.
            record.IsPublicHoliday = context.Calendar.IsPublicHoliday(record.AttendanceDate);
            record.IsValidDate = true;

            ProcessClockActivity(record, context);

            if (context.Calendar.IsDateInRange(record.AttendanceDate))
            {
                DispatchStatus(record, context);
            }

            // BUG FIX FROM PROJECT HISTORY: the day-of-week present count must
            // be updated UNCONDITIONALLY based on record.IsPresent, completely
            // independent of which status handler fired. In the original
            // buggy version, this increment lived INSIDE specific status
            // handlers (only "Present" and, for Consultants, "Single Punch
            // Absent"), which meant a day that was present AND a public
            // holiday, or present AND a weekend, silently never got counted
            // toward the Mon-Sun day tallies even though the person worked.
            if (record.IsPresent)
            {
                context.IncrementDayPresentCount(record.DayOfWeek);
            }

            ProcessWeekendOvertime(record, context);
        }
    }

    private static void ProcessClockActivity(DailyAttendanceRecord record, AttendanceContext context)
    {
        var employee = context.Employee;

        if (record.IsClockIn) employee.ClockInCount++;
        if (record.IsClockOut) employee.ClockOutCount++;

        if (record.IsPresent)
        {
            employee.PresentCount++;
            employee.ClockInOutCount++;
        }
    }

    /// <summary>
    /// Parses the raw status string into tokens, then walks PriorityLabels IN
    /// ORDER, dispatching to (and stopping at) the FIRST matching label. This
    /// single "return after first match" is what guarantees only one handler
    /// ever runs per day - the direct fix for the multi-status-token bug
    /// described in the PriorityLabels remarks above.
    /// </summary>
    private void DispatchStatus(DailyAttendanceRecord record, AttendanceContext context)
    {
        var tokens = AttendanceStatusParser.Parse(record.Status);

        foreach (var label in PriorityLabels)
        {
            if (tokens.Contains(label))
            {
                _handlers[label].Handle(record, context);
                return;
            }
        }
    }

    /// <summary>
    /// Tracks Saturday/Sunday overtime for Consultants only - Contractors have
    /// no weekend-OT concept under current policy (they're only ever evaluated
    /// against the public-holiday duration rule, handled separately in
    /// PublicHolidayStatusHandler).
    /// </summary>
    private static void ProcessWeekendOvertime(DailyAttendanceRecord record, AttendanceContext context)
    {
        var employee = context.Employee;

        if (employee.Type == ContractType.Contractor) return;
        if (!record.IsPresent || !record.IsWeekend) return;

        if (record.IsWorkDurationAtLeast(AppConfig.MinimumOtHours))
        {
            if (record.DayOfWeek == DayOfWeek.Saturday) employee.OtSaturdayCount++;
            else if (record.DayOfWeek == DayOfWeek.Sunday) employee.OtSundayCount++;
        }
    }
}
