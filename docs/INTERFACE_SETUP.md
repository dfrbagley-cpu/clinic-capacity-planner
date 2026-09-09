# Interface Version — local setup

The Interface Version is a Shiny application over the same R decision engine as
the Hospital R Tool. It is designed for one analyst on a local workstation. The
provided launcher listens only on `127.0.0.1`; it does not publish the application.

## Install and start

Install the reporting dependencies in `HOSPITAL_SETUP.md` and the `shiny` package
through your approved software process. No Node, Python, database or cloud
account is required. A browser displays the interface while R performs the work.

From the complete development folder, double-click `Run-Interface.cmd` or use:

```r
source("run_app.R")
# Or from a terminal: Rscript run_app.R
```

For the Windows launcher, set `CLINIC_RSCRIPT` to Rscript.exe if it is not on PATH.
Keep the R session/console open while using the interface. Stop R with Ctrl+C or
Escape when finished. Closing a browser tab does not necessarily stop R.

## Manager workflow

1. Choose Basic or Full. Select one or more original `.xlsx` visit exports.
2. Set the end date and number of weeks. Check Workbook options for password
   protection, worksheet number and text-date convention.
3. Select **Analyse my exports**. An import also creates the original complete
   workbook, so large exports may take a minute or more. Import settings take
   effect only when you analyse again.
4. Select a clinic. Current-activity cards cover all observed hours in the period.
   Choose a reference window and enter custom opening/closing times in HH:MM.
5. Review completed visits, distinct patients, cancellations, no-shows, added and
   excluded activity. Custom times recalculate using the R engine.
6. Download the **current comparison** for the active windows, or the
   **original report** for the saved windows plus staff and data-quality detail.
   Comparison downloads contain all clinics; the clinic picker filters the screen.

**Load fictional example** provides a safe demonstration. Its period is always
2026-07-06 through 2026-08-30, regardless of the import date controls. Full demo
mode loads the corresponding fictional planning workbook.

## Full mode

Download the Full input template after analysing your visits. Its keys and dates
come from that run. Fill in the relevant fields in Excel, select Full, attach it
alongside the visit exports and analyse again. Missing inputs remain missing.

Costed-hours recommendations refer to the **saved schedules** at import time.
The moving Custom window never acquires a cost estimate automatically. To cost a
new window, add it to `settings$schedules` in `config.R`, restart/reanalyse and
create a fresh planning template. Budget amounts use the currency of the supplied
workbook; the tool performs no currency conversion.

## Data handling and limits

Uploads are processed locally, in temporary session files. The app retains only
the necessary analysis fields in memory for distinct-patient calculations, plus
the original reporting workbook for download. Temporary session outputs are
removed on session close; an interrupted R process may leave temporary files for
normal workstation cleanup. Downloaded reports remain where you save them.
Shiny's file picker may retain uploaded temporary copies until the R process ends.

The app shows aggregate results. The original report still contains sensitive
provider, date and free-text detail. No uploaded records, passwords or reports
are committed or sent to a hosted service by this code. All interface assets are
served locally by Shiny. The upload limit is 200 MiB per request.

Password-protected inputs use the original Windows/Excel route. For other
systems, use approved unprotected `.xlsx` exports. Do not expose this launcher on
a network: hosted/multi-user authentication and deployment are future work.

The calculations have automated coverage. Browser interaction and protected
Windows Excel automation still need acceptance on the intended workstation.

The underlying local launch and sharing approach is documented by
[Posit](https://shiny.posit.co/r/articles/share/share/); application structure follows
[Shiny app formats](https://shiny.posit.co/r/articles/build/app-formats/).
