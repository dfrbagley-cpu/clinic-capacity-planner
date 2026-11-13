# Interactive fictional report

`clinic-capacity-fictional-demo.xlsx` is the actual Basic-mode output generated
by this project's R reporting pipeline from the 738 fictional rows in
`examples/synthetic_input/visits.xlsx`, with analysis end date **2026-08-30** and
an eight-week current period. It is not a hand-built mock-up. No real patient
records were used. Its Raw Data sheet omits MRN and Patient columns.

## Explore it

1. Download the XLSX and open it in Excel. GitHub's preview cannot run its formulas.
2. On **Compare Hours**, select Meridian or Cedar in the yellow Clinic cell.
3. Change the yellow Start and End times, including the Custom row. Enter normal
   Excel times such as `09:00` and `17:00`. Opening is included; closing excluded.
4. Compare completed visits, visits per recorded day, cancellations and no-shows.
   An end time at or before the start should display **Check times**.
5. Inspect **Schedule Comparison**, **Decision Evidence** and **Inputs and Rules**
   for the saved comparisons and assumptions.

The file has no macros, external workbook links or network data connections.
Keep Excel calculation on Automatic. It contains 101 native formulas and a
clinic dropdown. Security policies may open downloaded files in Protected View;
follow your organization's policy before editing, and never enable macros for
this example (none are present).

Custom hours recalculate visit counts, not unique-patient counts. Saved distinct
patient comparisons are on Schedule Comparison. Rerun the R tool or use the local
Shiny interface for distinct patients under new hours. More observed visits in a
window do not predict new demand or show that changing hours improves care.

## Provenance and checks

Generated with R 4.5.3 and openxlsx 4.2.9 from the source including PR #3.
Source fixture MD5: `627581f16831ff5fecdf9b20ab156883`.
No workbook content or styling was edited after pipeline generation.

Synthetic expected completed-visit counts for 08:00–16:00, 08:00–17:00 and
09:00–17:00 are Meridian **112 / 160 / 144**, Cedar **128 / 128 / 96**.
These agree with the R regression and workflow checks. Formula recalculation for
clinic changes, custom hours and an invalid window was independently checked in
a spreadsheet calculation engine. Desktop Excel was not available for a local
manual interaction check; hospital workstation and protected-input acceptance
remain separate. This sample is technical demonstration, not clinical validation.

To reproduce without changing the checked-in example:

```sh
Rscript run_report.R --mode basic --input examples/synthetic_input --output output/demo_basic --as-of 2026-08-30 --unprotected
```

Only refresh the public sample from the versioned fictional input. Never replace
it with a report generated from local clinical data.
