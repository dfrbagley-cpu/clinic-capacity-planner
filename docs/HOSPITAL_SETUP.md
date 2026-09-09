# Hospital R Tool — shared-drive setup

This guide covers the Hospital R Tool. The Interface Version has
its own launcher and an additional Shiny dependency; it is not required here.

## Runtime

Use R 4.1 or later (the original pipeline uses the native pipe). The development
check ran on R 4.3.3. Install these R packages once through the hospital's approved
software process: `dplyr`, `purrr`, `stringr`, `lubridate`, `openxlsx`, `tidyr`,
`tibble`, and their dependencies. `openxlsx` 4.2.5.2 was used in development.

Protected workbooks also require Windows desktop Excel and `excel.link` as in V3.
RStudio's `rstudioapi::askForPassword` provides the password prompt when using
RStudio. For the double-click launcher outside RStudio, install `getPass` for a
masked prompt. The program will not fall back to displaying a password as typed.

No Python, browser, cloud service, macro-enabled workbook, database, or internet
connection is needed during reporting. R runs on the analyst workstation; project,
input, and output files can reside on an accessible shared drive.

## Folder and path behaviour

- `input_data`: source exports only, never the generated report. Original headers
  begin at A1. Choose the correct worksheet using `settings$input_sheet`.
- `output`: each run gets its own folder, avoiding name collisions between users.
- `config.R`: input/output locations, Basic/Full mode, protected import setting,
  text-date convention, and analysis end date.
- `planning_inputs.xlsx`: optional Full inputs, never required by Basic.

Relative paths resolve from the project folder. In R strings, use forward slashes
or doubled backslashes. UNC examples: `//server/share/ClinicPlanner/input_data`.
The Windows launchers use `pushd` so UNC paths can be accessed by cmd.exe. Set
`CLINIC_RSCRIPT` to the actual IT-installed Rscript.exe path if it is not on PATH.
Users need read permission on code/inputs and write permission on the output
folder. Code should be maintained by an analyst; avoid concurrent edits to
`config.R` by giving analysts separate configuration copies if their settings differ.

The tool ignores Excel temporary files beginning with `~$`. Close source
workbooks before running protected import. Do not place a previous report among
the inputs. It rejects an identical input/output directory.

## Working with dates and hours

`as_of` is the last included day. The current period includes the previous 56
calendar days ending on that date; the prior period is the immediately preceding
56 days. All window comparisons use the same current period and comparison days.
Use completed reporting dates, not a partially finalised day.

Text dates use the explicit `mdy` or `dmy` choice in config. Excel date cells and
ISO dates are preferred. The legacy overview still includes all supplied records;
the new decision sheets use the documented analysis periods.

Add `settings$schedules` in config for custom saved windows and patient counts:

```r
settings$schedules <- data.frame(
  option = c("8-4", "8-5", "9-5", "8:30-4:30"),
  start = c(8, 8, 9, 8.5) * 3600,
  end = c(16, 17, 17, 16.5) * 3600
)
```

The first row is the comparison reference, not a statement that those are the
clinic's actual current operating hours. Overnight windows are not supported.
Excel start/end inputs accept times from 00:00 through an end of 24:00.

## Output handling

The output is a macro-free `.xlsx` with normal formulas. Use automatic calculation
in Excel. Spreadsheet previews that do not calculate formulas may show blank
formula cells until the file is opened in Excel. Static Schedule Comparison
already contains the configured-window results.

The original workbook output is not encrypted by this tool. Omitting two patient
identifier columns does not make provider/date/free-text detail anonymous. Apply
the hospital's normal folder permissions and report-handling practices. Passwords
are prompted at runtime and not written to the report or a config file.

Run the workstation acceptance checklist before routine use; network-drive
permissions and protected Excel COM behaviour cannot be verified on Linux.
