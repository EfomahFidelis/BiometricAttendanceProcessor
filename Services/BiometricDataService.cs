using ClosedXML.Excel;
using BiometricAttendanceProcessor.Configuration;
using BiometricAttendanceProcessor.Models;

namespace BiometricAttendanceProcessor.Services;

/// <summary>
/// Reads the raw "DataSheet" worksheet and extracts each employee's slice of
/// attendance rows for the current payroll cycle.
/// </summary>
public interface IBiometricDataService
{
    List<BiometricRecord> LoadAll();
    int FindFirstRowIndexForEmployee(List<BiometricRecord> allRecords, string employeeId);
    DateTime GetFirstDateForEmployee(List<BiometricRecord> allRecords, string employeeId, int startIndex);
    List<BiometricRecord> ExtractEmployeeAttendance(
        List<BiometricRecord> allRecords, string employeeId, int startIndex, int daysToExtract, bool isIrregular);
}

/// <summary>
/// Another REPOSITORY-pattern class (see EmployeeRepository for the full
/// explanation) - this one owns "how do we get biometric attendance data",
/// isolating every other class from the fact that today, that means reading an
/// Excel worksheet.
/// </summary>
public class BiometricDataService : IBiometricDataService
{
    private readonly IXLWorksheet _sheet;

    public BiometricDataService(IXLWorksheet dataSheet)
    {
        _sheet = dataSheet;
    }

    /// <summary>
    /// Loads the entire DataSheet into memory once, as a flat list. Loading
    /// everything up-front (rather than re-querying the worksheet per employee,
    /// as the VBA version effectively did via repeated Range lookups) is both
    /// simpler to reason about and much faster for a worksheet with thousands
    /// of rows, since we only touch each Excel cell exactly once.
    /// </summary>
    public List<BiometricRecord> LoadAll()
    {
        var records = new List<BiometricRecord>();
        int lastRow = _sheet.LastRowUsed()?.RowNumber() ?? 1;

        for (int row = 2; row <= lastRow; row++)
        {
            var idCell = _sheet.Cell(row, (int)BiometricColumn.EmployeeId);
            if (idCell.IsEmpty()) continue;

            records.Add(new BiometricRecord
            {
                EmployeeId = idCell.GetString().Trim(),
                Date = _sheet.Cell(row, (int)BiometricColumn.Date).GetDateTime(),
                ClockIn = ReadTime(_sheet.Cell(row, (int)BiometricColumn.ClockIn)),
                ClockOut = ReadTime(_sheet.Cell(row, (int)BiometricColumn.ClockOut)),
                Status = _sheet.Cell(row, (int)BiometricColumn.Status).GetString(),
                TotalWorkDurationHours = ReadDuration(_sheet.Cell(row, (int)BiometricColumn.TotalWorkDuration)),
                SourceRowNumber = row
            });
        }

        return records;
    }

    /// <summary>
    /// Reads a clock-in/clock-out cell as a clean, unambiguous TimeSpan?.
    ///
    /// THIS IS THE C# FIX FOR THE VBA "IsDate() ALWAYS RETURNS FALSE" BUG we
    /// spent considerable effort diagnosing. That bug existed because VBA stored
    /// the time as a bare Variant/Double and IsDate() refused to recognize a raw
    /// fractional number as a valid time. Here, we settle the question exactly
    /// ONCE, at the moment of reading the cell, and from then on the rest of the
    /// program only ever sees a clean `TimeSpan?` - either it HasValue (a real
    /// time) or it doesn't (blank / "N.A." / unparseable). No other class ever
    /// needs to re-derive "is this actually a valid time?" again.
    /// </summary>
    private static TimeSpan? ReadTime(IXLCell cell)
    {
        if (cell.IsEmpty()) return null;

        var text = cell.GetString().Trim();
        if (text.Length == 0 || text == "N.A.") return null;

        // Excel time cells are stored as fractional days internally, whether
        // formatted as "Custom" or as a real Time format - ClosedXML surfaces
        // this either as a DateTime (whose TimeOfDay we want) or as a raw
        // double fraction, depending on formatting. We handle both so the
        // result is correct regardless of how the source cell was formatted.
        if (cell.TryGetValue(out DateTime dateTimeValue))
            return dateTimeValue.TimeOfDay;

        if (cell.TryGetValue(out double fractionOfDay))
            return TimeSpan.FromDays(fractionOfDay);

        return null;
    }

    /// <summary>
    /// Reads a Total Work Duration cell (Column J) as hours.
    ///
    /// Deliberately does NOT go through DateTime.TimeOfDay, unlike ReadTime:
    /// TimeOfDay wraps at 24 hours and would silently truncate a duration like
    /// 24:01:26 down to 00:01:26. Reading the raw fraction-of-a-day double and
    /// converting with TimeSpan.FromDays preserves durations greater than 24
    /// hours, which is exactly what overnight "Shift Change" shifts produce.
    /// </summary>
    private static double? ReadDuration(IXLCell cell)
    {
        if (cell.IsEmpty()) return null;

        var text = cell.GetString().Trim();
        if (text.Length == 0 || text == "N.A.") return null;

        // Excel stores a duration as a fractional number of days (1.0 = 24h).
        // This branch covers both cells formatted as a Time and cells formatted
        // as a raw decimal - the underlying value is identical either way.
        if (cell.TryGetValue(out double fractionOfDay))
            return TimeSpan.FromDays(fractionOfDay).TotalHours;

        // Safety net: some exports write the duration as a literal "hh:mm:ss"
        // string rather than a real Excel time value.
        if (TimeSpan.TryParse(text, out var parsed))
            return parsed.TotalHours;

        return null;
    }

    public int FindFirstRowIndexForEmployee(List<BiometricRecord> allRecords, string employeeId)
    {
        for (int i = 0; i < allRecords.Count; i++)
        {
            if (allRecords[i].EmployeeId == employeeId) return i;
        }
        return -1;
    }

    public DateTime GetFirstDateForEmployee(List<BiometricRecord> allRecords, string employeeId, int startIndex)
    {
        for (int i = startIndex; i < allRecords.Count; i++)
        {
            if (allRecords[i].EmployeeId == employeeId) return allRecords[i].Date;
        }
        return DateTime.Today;
    }

    /// <summary>
    /// Extracts exactly this employee's attendance rows for the current payroll
    /// cycle, starting from startIndex in the master list.
    ///
    /// THIS IS THE C# VERSION OF THE FIX WE MADE FOR THE VBA EXTRACTION BUG.
    /// Originally (in the "Class Approach"), the very first matching row was
    /// ALWAYS discarded, regardless of whether it was a genuine payroll-cycle
    /// start date or a stray leftover day from the previous cycle. The fix -
    /// preserved faithfully here - is that the first match is only discarded
    /// when `isIrregular` is true, i.e. when the biometric export actually did
    /// start one day early (on the 18th instead of the 19th).
    /// </summary>
    public List<BiometricRecord> ExtractEmployeeAttendance(
        List<BiometricRecord> allRecords, string employeeId, int startIndex, int daysToExtract, bool isIrregular)
    {
        var extracted = new List<BiometricRecord>();

        // When NOT irregular, there is no stray day to discard, so we treat the
        // "first match" as already handled and go straight to collecting rows.
        bool firstMatchDiscarded = !isIrregular;

        bool everFoundMatch = false;
        int notFoundStreak = 0;

        for (int i = startIndex; i < allRecords.Count; i++)
        {
            if (allRecords[i].EmployeeId == employeeId)
            {
                notFoundStreak = 0;

                if (!firstMatchDiscarded)
                {
                    // This row is the stray leftover day from the previous
                    // payroll cycle - copy-then-discard, matching Direct
                    // Approach's behaviour exactly.
                    firstMatchDiscarded = true;
                    continue;
                }

                extracted.Add(allRecords[i]);
                everFoundMatch = true;

                if (extracted.Count == daysToExtract) break;
            }
            else if (everFoundMatch)
            {
                // We've already started collecting this employee's rows, but
                // this row belongs to someone else. Allow up to 10 consecutive
                // "someone else" rows before giving up - the source data is
                // grouped by employee but isn't guaranteed to be perfectly
                // contiguous, so a small tolerance avoids stopping early on a
                // single stray interleaved row.
                notFoundStreak++;
                if (notFoundStreak >= 10) break;
            }
        }

        return extracted;
    }
}