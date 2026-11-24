using ClosedXML.Excel;
using BiometricAttendanceProcessor.Analysis;
using BiometricAttendanceProcessor.Configuration;
using BiometricAttendanceProcessor.Models;
using BiometricAttendanceProcessor.Services;

namespace BiometricAttendanceProcessor.Reporting;

/// <summary>
/// Writes one row per employee to the "Extraction Records" report worksheet,
/// including the full day-by-day attendance code grid with colour-coding.
/// Equivalent of the VBA "CAttendanceReportGenerator" class.
/// </summary>
public class AttendanceReportGenerator
{
    /// <summary>
    /// Colour lookup for each attendance code, mirroring the VBA
    /// FormatGridCell Select Case block. Kept as a simple static Dictionary
    /// here rather than one big switch statement, purely for readability -
    /// adding a new code+colour pair is a one-line change.
    /// </summary>
private static readonly Dictionary<string, (XLColor Back, XLColor Font)> CodeColors = new()
{
    ["P"]    = (XLColor.FromArgb(0, 176, 80),    XLColor.White),
    ["W/O"]  = (XLColor.FromArgb(64, 64, 64),    XLColor.Black),
    ["T/O"]  = (XLColor.FromArgb(64, 64, 64),    XLColor.Black),
    ["XWO"]  = (XLColor.FromArgb(175, 175, 175), XLColor.Black),
    ["L"]    = (XLColor.FromArgb(64, 64, 64),    XLColor.Black),
    ["PL"]   = (XLColor.FromArgb(0, 128, 128),   XLColor.Black),
    ["SPA"]  = (XLColor.FromArgb(0, 112, 192),   XLColor.White),
    ["uL"]   = (XLColor.FromArgb(128, 128, 128), XLColor.Black),
    ["PH"]   = (XLColor.FromArgb(204, 0, 102),   XLColor.White),
    ["PhP"]  = (XLColor.FromArgb(204, 0, 102),   XLColor.White),
    ["XPH"]  = (XLColor.FromArgb(64, 64, 64),    XLColor.White),
    ["XTO"]  = (XLColor.FromArgb(64, 64, 64),    XLColor.Black),  // not in VBA; choose Black
    ["pSPA"] = (XLColor.FromArgb(102, 153, 255), XLColor.White),
    ["3HL"]  = (XLColor.FromArgb(255, 153, 0),   XLColor.White),
    ["X"]    = (XLColor.Yellow,                  XLColor.Black),
    ["eXP"]  = (XLColor.Red,                     XLColor.Black),
    [""]     = (XLColor.Black,                   XLColor.White),
};

    private readonly IXLWorksheet _sheet;
    private int _currentRow = 4; // Row 4 is the first data row; rows 1-3 are headers.

    public AttendanceReportGenerator(IXLWorksheet reportSheet)
    {
        _sheet = reportSheet;
        WriteHeaders();
    }

    /// <summary>
    /// EXTENSION POINT: this writes only the minimal set of column headers
    /// needed to make the report readable. The original VBA report has a much
    /// richer header layout (merged cells, the day-of-month header row above
    /// the attendance grid, coloured section headers, etc.) - if you want to
    /// reproduce that exactly, this is the method to expand. It's deliberately
    /// kept simple here so the underlying data logic isn't obscured by
    /// formatting code.
    /// </summary>
    private void WriteHeaders()
    {
        var headers = new (ReportColumn Column, string Text)[]
        {
            (ReportColumn.EmployeeId, "Staff No"),
            (ReportColumn.EmployeeName, "Name"),
            (ReportColumn.ContractType, "Contract Type"),
            (ReportColumn.Mandays, "Month Days"),
            (ReportColumn.Present, "Present"),
            (ReportColumn.MissingDays, "Missing/Blank Days"),
            (ReportColumn.ClockIns, "Clock-Ins"),
            (ReportColumn.ClockOuts, "Clock-Outs"),
            (ReportColumn.ClockInOuts, "Clocked In & Out"),
            (ReportColumn.Absent, "Absent"),
            (ReportColumn.PresentOnLeave, "Present on Leave"),
            (ReportColumn.Leave, "Leave/Unpaid Leave"),
            (ReportColumn.WeeklyOff, "Weekly Off/Time Off"),
            (ReportColumn.SunPresent, "Sun"),
            (ReportColumn.MonPresent, "Mon"),
            (ReportColumn.TuePresent, "Tue"),
            (ReportColumn.WedPresent, "Wed"),
            (ReportColumn.ThuPresent, "Thu"),
            (ReportColumn.FriPresent, "Fri"),
            (ReportColumn.SatPresent, "Sat"),
            (ReportColumn.OtSunday, "OT Sun"),
            (ReportColumn.OtSaturday, "OT Sat"),
            (ReportColumn.PublicHolidayCount, "Total Public Holiday"),
        };

        foreach (var (column, text) in headers)
        {
            var cell = _sheet.Cell(3, (int)column);
            cell.Value = text;
            cell.Style.Font.Bold = true;
        }

        // Named-holiday sub-headers, one per configured holiday, positioned
        // relative to PublicHolidayCount exactly as AttendanceContext's
        // per-holiday counts are written below.
        for (int i = 0; i < AppConfig.PublicHolidays.Length; i++)
        {
            var cell = _sheet.Cell(3, (int)ReportColumn.PublicHolidayCount + i + 1);
            cell.Value = AppConfig.PublicHolidays[i].ToString("dd-MMM");
            cell.Style.Font.Bold = true;
        }
    }

    /// <summary>Writes a single row for an employee who has no valid ID or was not found in the biometric data.</summary>
    public void WriteEmployeeNotFound(Employee employee)
    {
        _sheet.Cell(_currentRow, (int)ReportColumn.EmployeeId).Value = employee.EmployeeId;
        _sheet.Cell(_currentRow, (int)ReportColumn.EmployeeName).Value = employee.Name;
        _sheet.Cell(_currentRow, (int)ReportColumn.ContractType).Value = employee.Type.ToString();
        _currentRow++;
    }

    /// <summary>Writes a full attendance summary row, including the day-by-day code grid.</summary>
    public void WriteEmployeeAttendance(Employee employee, AttendanceContext context, PayrollCalendar calendar)
    {
        int row = _currentRow;

        WriteSummaryColumns(row, employee, context);
        WriteHolidayColumns(row, context, calendar);
        int missingDays = WriteAttendanceGrid(row, context, calendar);

        employee.MissingBlankDaysCount = missingDays;
        _sheet.Cell(row, (int)ReportColumn.MissingDays).Value = missingDays;

        _currentRow++;
    }

    private void WriteSummaryColumns(int row, Employee employee, AttendanceContext context)
    {
        _sheet.Cell(row, (int)ReportColumn.EmployeeId).Value = employee.EmployeeId;
        _sheet.Cell(row, (int)ReportColumn.EmployeeName).Value = employee.Name;
        _sheet.Cell(row, (int)ReportColumn.ContractType).Value = employee.Type.ToString();
        _sheet.Cell(row, (int)ReportColumn.Mandays).Value = employee.WorkingDaysInMonth;
        _sheet.Cell(row, (int)ReportColumn.Present).Value = employee.PresentCount;
        _sheet.Cell(row, (int)ReportColumn.ClockIns).Value = employee.ClockInCount;
        _sheet.Cell(row, (int)ReportColumn.ClockOuts).Value = employee.ClockOutCount;
        _sheet.Cell(row, (int)ReportColumn.ClockInOuts).Value = employee.ClockInOutCount;
        _sheet.Cell(row, (int)ReportColumn.Absent).Value = employee.AbsentCount;
        _sheet.Cell(row, (int)ReportColumn.PresentOnLeave).Value = employee.WorkingOnLeaveCount;
        _sheet.Cell(row, (int)ReportColumn.Leave).Value = employee.LeaveCount;
        _sheet.Cell(row, (int)ReportColumn.WeeklyOff).Value = employee.WeeklyOffCount;

        // DayPresentCounts is indexed 0=Sunday..6=Saturday; ReportColumn.SunPresent
        // through SatPresent are laid out in the same Sun-first order, so a
        // straight loop with no re-mapping is needed.
        for (int i = 0; i < 7; i++)
        {
            _sheet.Cell(row, (int)ReportColumn.SunPresent + i).Value = context.DayPresentCounts[i];
        }

        _sheet.Cell(row, (int)ReportColumn.OtSunday).Value = employee.OtSundayCount;
        _sheet.Cell(row, (int)ReportColumn.OtSaturday).Value = employee.OtSaturdayCount;
    }

    private void WriteHolidayColumns(int row, AttendanceContext context, PayrollCalendar calendar)
    {
        _sheet.Cell(row, (int)ReportColumn.PublicHolidayCount).Value = context.PublicHolidayCounts.Count;

        foreach (var (date, count) in context.PublicHolidayCounts)
        {
            int holidayIndex = calendar.GetPublicHolidayIndex(date);
            if (holidayIndex > 0)
            {
                _sheet.Cell(row, (int)ReportColumn.PublicHolidayCount + holidayIndex).Value = count;
            }
        }
    }

    /// <summary>Writes the day-by-day colour-coded attendance grid. Returns the count of missing/blank days.</summary>
    private int WriteAttendanceGrid(int row, AttendanceContext context, PayrollCalendar calendar)
    {
        int dayColumn = 0;
        int missingDays = 0;

        foreach (var date in calendar.GetPayrollDates())
        {
            var code = context.GetCode(date);
            if (string.IsNullOrEmpty(code)) missingDays++;

            // Day 1's code starts one column AFTER AttendanceGridStart, to line up
            // with the day-number/day-name/date header row (which begins at AF).
            var cell = _sheet.Cell(row, (int)ReportColumn.AttendanceGridStart + dayColumn + 1);

            if (CodeColors.TryGetValue(code, out var colours))
            {
                cell.Value = code;
                cell.Style.Fill.BackgroundColor = colours.Back;
                cell.Style.Font.FontColor = colours.Font;
            }
            else
            {
                // Matches VBA's "Case Else": unrecognised code is written as "ER"
                // on a white background with black text.
                cell.Value = "ER";
                cell.Style.Fill.BackgroundColor = XLColor.White;
                cell.Style.Font.FontColor = XLColor.Black;
            }

            cell.Style.Font.Bold = true;
            dayColumn++;
        }

        return missingDays;
    }
}
