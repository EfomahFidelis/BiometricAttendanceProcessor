using BiometricAttendanceProcessor.Models;

namespace BiometricAttendanceProcessor.Analysis;

/// <summary>
/// DESIGN PATTERN: STRATEGY.
///
/// Each concrete implementation of this interface (see the StatusHandlers/
/// folder) knows how to process exactly ONE kind of attendance status - Present,
/// Absent, Leave, Public Holiday, etc. AttendanceAnalyzer never contains an
/// "if status is this, do that; if status is that, do this" chain of logic
/// itself - it just looks up the right IAttendanceStatusHandler for the winning
/// status label and calls Handle() on it.
///
/// Why this matters: adding a brand-new attendance status in the future (say,
/// "Maternity Leave") means writing ONE new class that implements this
/// interface and registering it in AttendanceAnalyzer's constructor. You do not
/// need to touch, re-read, or risk breaking any of the existing handlers - each
/// one is fully independent and individually testable.
/// </summary>
public interface IAttendanceStatusHandler
{
    void Handle(DailyAttendanceRecord record, AttendanceContext context);
}

/// <summary>
/// Shared helper for splitting a raw status string like
/// "Present,Holiday (Eid el Fitri),Request" into clean tokens:
/// ["Present", "Holiday", "Request"] - stripping any parenthesized detail text
/// and surrounding whitespace. Equivalent of VBA's ParseStatusString/CleanStatusPart.
/// </summary>
public static class AttendanceStatusParser
{
    public static List<string> Parse(string status)
    {
        if (string.IsNullOrWhiteSpace(status)) return new List<string>();

        return status
            .Split(',')
            .Select(CleanToken)
            .ToList();
    }

    private static string CleanToken(string token)
    {
        token = token.Trim();
        int bracketPos = token.IndexOf('(');
        if (bracketPos > 0)
        {
            token = token[..bracketPos].Trim();
        }
        return token;
    }
}
