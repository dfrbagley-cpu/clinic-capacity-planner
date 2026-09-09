# Clinic Capacity Planner — Hospital R Tool

This folder can live on your hospital shared drive. R runs on the analyst's
computer and reads the existing Excel exports. Managers receive an Excel report.

1. **One-time setup:** ask IT to install R and the reporting packages listed in
   `docs/HOSPITAL_SETUP.md`. Password-protected exports also need Windows desktop
   Excel and `excel.link`, as in the original tool.
2. **Copy the whole folder** to an approved shared-drive location. Keep the folder
   structure intact. You do not need GitHub or the interface package to run it.
3. **Put visit exports in `input_data`.** Use the original 12 column headers.
4. **Edit `config.R`:** set the analysis end date, input/output folders and whether
   the workbooks are password protected. Dates may use the configured text-date
   order. Do not write a workbook password into this file.
5. **Double-click `Run-Basic.cmd`.** IT can set `CLINIC_RSCRIPT` if Rscript is not
   on PATH. Alternatively, open the RStudio project and run `source("run_report.R")`.
6. **Open the new workbook in `output/run_...`.** Compare hours on the first sheet.
   Basic needs only the visit export. Full uses the same launcher workflow with
   `Run-Full.cmd` and an additional planning workbook.

The saved comparisons include distinct patients. Excel's editable time cells
recalculate visit counts; rerun R for distinct patients in newly configured
windows. The Interface Version can recalculate those patient counts live.

First use: reconcile a completed period with the old report. Windows protected
import and shared-drive permissions still need the acceptance check in
`docs/ACCEPTANCE.md` on a hospital workstation.

R and packages are installed once. There are no online services or updates during
normal reporting. Keep generated reports in approved hospital folders.
