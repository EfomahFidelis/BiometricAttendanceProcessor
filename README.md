# Biometric Attendance Processor (C# / .NET 9)

This is the C# port of the VBA "Class Approach" biometric attendance system,
carrying forward every bug fix worked out during the VBA debugging process:
extraction chop-off logic, single-status-per-day priority dispatch, the
`IsDate()`/Variant duration bug, decoupled day-present counting, named-holiday
breakdown, and the Consultant/Contractor absenteeism rules (including the
`XWO` vs `W/O` distinction and the defensive `XTO` handling).

## 1. Project structure

```
BiometricAttendanceProcessor/
├── Configuration/
│   ├── AppConfig.cs          <- payroll dates, holidays, OT threshold (EDIT MONTHLY)
│   └── ColumnMappings.cs     <- worksheet column layout enums
├── Models/
│   ├── Employee.cs           <- employee entity + ContractType enum
│   ├── DailyAttendanceRecord.cs  <- one analyzed day of attendance
│   └── BiometricRecord.cs    <- one raw row from DataSheet
├── Services/
│   ├── PayrollCalendar.cs
│   ├── EmployeeRepository.cs     <- reads CriteriaSheet (Repository pattern)
│   ├── BiometricDataService.cs   <- reads DataSheet + extraction algorithm
│   └── FileSystemService.cs
├── Analysis/
│   ├── AttendanceContext.cs      <- per-employee working state
│   ├── IAttendanceStatusHandler.cs   <- Strategy interface
│   ├── StatusHandlers/           <- one class per attendance status
│   └── AttendanceAnalyzer.cs     <- priority dispatch + orchestration
├── Reporting/
│   └── AttendanceReportGenerator.cs  <- writes "Extraction Records" sheet
├── Processing/
│   └── AttendanceProcessor.cs    <- Facade tying everything together
└── Program.cs                    <- composition root / entry point
```

**Design patterns used, and why** (each is also called out in the code comments
at its point of use):

| Pattern | Where | Why |
|---|---|---|
| **Repository** | `EmployeeRepository`, `BiometricDataService` | Hides "the data lives in an Excel sheet" behind an interface, so the storage mechanism can change without touching business logic. |
| **Strategy** | `IAttendanceStatusHandler` + `StatusHandlers/*` | Each attendance status (Present, Absent, Leave, ...) is its own interchangeable class. Adding a new status = adding one new class, with zero risk to existing ones. |
| **Facade** | `AttendanceProcessor` | Hides the coordination of six collaborating classes behind one `ProcessAll()` call. |
| **Dependency Injection (manual, via constructors)** | Everywhere; wired in `Program.cs` | Every class depends on interfaces/abstractions passed into its constructor, never on concretions it creates itself - this is what makes each piece independently testable and replaceable. |

## 2. Prerequisites

You already have what's needed:

```
dotnet --version        -> 9.0.317
dotnet --list-runtimes   -> Microsoft.NETCore.App 9.0.19 present
```

In VS Code, install these two extensions (one-time setup):
1. **C# Dev Kit** (by Microsoft) - search "C# Dev Kit" in the Extensions panel (`Ctrl+Shift+X`).
2. It will automatically pull in the base **C#** extension and **.NET Install Tool** as dependencies.

## 3. Opening the project in VS Code

1. Extract/save this whole `BiometricAttendanceProcessor` folder somewhere on disk, e.g. `C:\Dev\BiometricAttendanceProcessor`.
2. In VS Code: **File → Open Folder...** and select that folder (the one containing `BiometricAttendanceProcessor.csproj`).
3. VS Code will detect the C# project and may prompt "Required assets to build and debug are missing. Add them?" - click **Yes**. This generates a `.vscode/launch.json` and `.vscode/tasks.json` for you automatically.

## 4. Restoring packages and building

Open a terminal inside VS Code (**Terminal → New Terminal**, or `` Ctrl+` ``) and run:

```powershell
dotnet restore
dotnet build
```

`dotnet restore` downloads the `ClosedXML` NuGet package referenced in the
`.csproj` file (needs an internet connection the first time). `dotnet build`
compiles the project - you should see `Build succeeded.` with 0 errors.

## 5. Preparing your input file

The program expects a single `.xlsx` workbook containing two worksheets named
exactly:
- `CriteriaSheet` (employee master data - same layout as the VBA version)
- `DataSheet` (raw biometric export - same layout as the VBA version)

Any existing `Extraction Records` sheet in that same workbook will be cleared
and reused; if it doesn't exist, it will be created.

## 6. Running the program

### Option A - from the terminal, with arguments

```powershell
dotnet run -- "C:\Data\PayrollSource.xlsx" "C:\Data\Output"
```

The two arguments are: path to the source workbook, and the folder where the
report and per-employee extraction files should be written (it will be
created if it doesn't exist).

### Option B - from the terminal, interactively

```powershell
dotnet run
```

The program will prompt you for both paths if you don't supply them as
arguments.

### Option C - with the debugger (recommended while learning/modifying the code)

1. Open `Program.cs`.
2. Click in the left margin next to any line (e.g. inside `Main`) to set a
   **breakpoint** (a red dot appears).
3. Press **F5** (or the "Run and Debug" icon in the sidebar, then the green
   play arrow).
4. The first time, VS Code may ask you to select an environment - choose
   **C#**.
5. Because the program reads arguments/console input, it will pause and let
   you type the source path and output folder directly into the **Debug
   Console / Terminal** panel that opens.
6. Execution will stop at your breakpoint, and you can hover over variables,
   step through code line-by-line (`F10` = step over, `F11` = step into), and
   watch exactly how a specific employee's record flows through
   `AttendanceAnalyzer.Analyze()` - this is the single most useful thing you
   can do to understand how the priority dispatch and status handlers work
   together.

To supply the two path arguments when debugging (instead of typing them at
the prompt), edit `.vscode/launch.json` and add an `"args"` entry:

```jsonc
{
    "configurations": [
        {
            "name": "C#: BiometricAttendanceProcessor Debug",
            // ...existing settings...
            "args": ["C:\\Data\\PayrollSource.xlsx", "C:\\Data\\Output"]
        }
    ]
}
```

## 7. What gets produced

After a run, in your chosen output folder you'll find:
- `AttendanceReport.xlsx` - the source workbook, saved with the populated
  "Extraction Records" sheet.
- A timestamped subfolder (e.g. `2026-04-19-14-30-05\`) containing one `.xlsx`
  file per employee with their extracted, payroll-cycle-aligned raw attendance
  rows - the audit trail equivalent of the VBA `SaveExtractedBiometricAttendance` output.

## 8. Segments you'll likely want to modify

These are marked with `EXTENSION POINT` comments in the code itself:

- **`AttendanceReportGenerator.WriteHeaders()`** - currently writes a minimal
  header row. The original VBA report has richer merged-cell headers and
  colour-coded section bands; expand this method to match if you want that.
- **`Program.Main()`** - the VBA version prompted to sort the report by
  multiple columns before saving. That's not implemented here; `IXLRange.Sort`
  (via `SortFields.Add`) is the ClosedXML equivalent if you want to add it.
- **`AppConfig`** - update `PayrollStartDate`, `PayrollEndDate`, and
  `PublicHolidays` every payroll cycle. This is the *only* file that should
  need routine, non-programmer-level edits.
- **Adding a new attendance status** - implement `IAttendanceStatusHandler`,
  register it in `AttendanceAnalyzer`'s constructor's `_handlers` dictionary,
  and add its label to `PriorityLabels` at the correct priority position. No
  other file needs to change.

## 9. A note on testing your changes

Because every service is exposed through an interface
(`IEmployeeRepository`, `IBiometricDataService`), you can write unit tests
(e.g. with **xUnit**, added via `dotnet add package xunit`) that construct an
`AttendanceAnalyzer` directly and feed it hand-built `DailyAttendanceRecord`
lists, with no Excel file involved at all. That's a natural next step once
you're comfortable with the current structure - it would have caught several
of the bugs we found in the VBA version (like the duplicated `"3HL3HL"` codes)
immediately, rather than only after inspecting a full report output.
