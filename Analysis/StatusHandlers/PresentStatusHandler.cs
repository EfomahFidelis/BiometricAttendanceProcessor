using BiometricAttendanceProcessor.Models;

namespace BiometricAttendanceProcessor.Analysis.StatusHandlers;

/// <summary>
/// Handles the "Present" status token. Notice this handler DEPENDS ON two other
/// strategies (PublicHolidayStatusHandler and AbsentStatusHandler) and delegates
/// to them under certain conditions, rather than duplicating their logic. This
/// composition is wired up once, in AttendanceAnalyzer's constructor - see the
/// comment there for why that's the "composition root" for this small object graph.
/// </summary>
public class PresentStatusHandler : IAttendanceStatusHandler
{
    private readonly PublicHolidayStatusHandler _publicHolidayHandler;
    private readonly AbsentStatusHandler _absentHandler;

    public PresentStatusHandler(PublicHolidayStatusHandler publicHolidayHandler, AbsentStatusHandler absentHandler)
    {
        _publicHolidayHandler = publicHolidayHandler;
        _absentHandler = absentHandler;
    }

    public void Handle(DailyAttendanceRecord record, AttendanceContext context)
    {
        // A day tagged "Present" that also happens to be a public holiday needs
        // the fuller holiday-specific logic (duration threshold, SPA handling,
        // etc.), so we redirect there rather than just stamping a plain "P".
        if (record.IsPublicHoliday)
        {
            _publicHolidayHandler.Handle(record, context);
            return;
        }

        if (record.IsPresent)
        {
            context.StoreCode(record.AttendanceDate, "P");
        }
        else
        {
            // Status text says "Present" but the clock data disagrees (no valid
            // clock-in/clock-out pair) - fall through to the Absent rules so the
            // correct contract-type-specific code and counters are applied.
            _absentHandler.Handle(record, context);
        }
    }
}
