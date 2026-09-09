# Clinic Capacity Planner

**Compare clinic hours using the visits you already record.**

A decision-support tool for clinic managers and hospital analysts, delivered as
the **Hospital R Tool** and the **Interface Version**. Both use the same R engine.
The first question is: **if recorded visit patterns stayed the same, what activity
would fall within 8–4, 8–5, or 9–5?**

This public project shows the development of an existing visit-reporting workflow
into operational decision support: defining useful measures, handling data quality,
comparing service options and making the assumptions behind a recommendation
inspectable. All included patient records are fictional.

**Explore the project:** [design decisions and project overview](docs/PROJECT_OVERVIEW.md)
· [run the interface](docs/INTERFACE_SETUP.md)
· [hospital setup](START-HOSPITAL.md)
· [verification and remaining acceptance work](docs/ACCEPTANCE.md)
· [license](LICENSE).

## A concrete example

The supplied fictional dataset covers two clinics. These are completed visits
captured during the current eight-week reporting period, with recorded visit
times held constant:

| Clinic | 08:00–16:00 | 08:00–17:00 | 09:00–17:00 |
|---|---:|---:|---:|
| Meridian Clinic | 112 | 160 | 144 |
| Cedar Clinic | 128 | 128 | 96 |

For Meridian, shifting an eight-hour window from 08:00–16:00 to 09:00–17:00
captures 32 more completed visits. For Cedar, that shift excludes 32 completed
visits. The same schedule change therefore has different implications by clinic.
Neither comparison predicts how patients would reschedule or proves that a
change would improve access. Full mode adds explicit staffing and budget inputs.

To reproduce this example, start the Interface Version and select **Load fictional
example**, or run `Rscript examples/run_demo.R` after installing the reporting
dependencies. See [the original-schema fictional workbook](examples/synthetic_input/visits.xlsx).

## Two versions, one calculation engine

| Delivery version | How to use it | Requirements |
|---|---|---|
| **Hospital R Tool** | Copy the release folder to the shared drive, set `config.R`, run `Run-Basic.cmd` or `Run-Full.cmd`, open the Excel report | R and the reporting packages; no Shiny or hosting |
| **Interface Version** | Run `Run-Interface.cmd` or `source("run_app.R")`; load exports, select a clinic, edit hours and download results | The same R packages plus `shiny`; a local browser |

The source for both versions is maintained together in the public
[`clinic-capacity-planner` repository](https://github.com/dfrbagley-cpu/clinic-capacity-planner).
This avoids two implementations drifting apart. The applications run locally;
this repository publishes the source and fictional examples.

The interface recalculates **distinct patients as well as visits** for custom
hours. It reads the same Excel export and offers the same Basic and Full modes.
See [interface setup](docs/INTERFACE_SETUP.md).

## Two analysis modes within either version

| | Basic | Full |
|---|---|---|
| Input | The same 12 columns used by the original Visit Analysis Tool V3 | Those visits plus an optional planning workbook |
| Main decision | Compare recorded activity within alternative opening hours | Compare costed opening-hour options within a budget, plus constrained hours transfers |
| Measures | Completed visits, distinct patients, cancellations, no-shows, group attendances and sessions | Also completed visits per staffed hour, allocated cost per completed visit, offered-capacity use |
| Recommendations | Explainable investigation and pilot suggestions | Conditional care/cost estimates with demand, budget and access checks |
| Hospital operation | Shared-drive folders, R on the analyst workstation, Excel output | The same workflow |
| Required external services | None: no cloud account, external API, database or internet at runtime | None |

Basic never requires Full inputs. Missing or invalid Full inputs are reported;
the existing activity and Basic recommendations remain available.

## Start the Hospital R Tool

1. Put the project folder on the hospital shared drive.
2. Have IT install R and the packages in [the hospital setup guide](docs/HOSPITAL_SETUP.md).
3. Put the existing monthly Excel exports in `input_data`.
4. Set the analysis end date and export options in `config.R`.
5. Open `Clinic-Capacity-Planner.Rproj` in RStudio and run `source("run_report.R")`.
   Alternatively, use `Run-Basic.cmd` after IT configures Rscript and the masked
   password prompt. The original exports can remain password protected.
6. Open the workbook under `output/run_...`. Managers can use the report in Excel
   without R. R is needed again when new exports arrive or saved patient-count
   windows change.

The report opens on **Compare Hours**. Select a clinic and edit the yellow start
and end cells. Native Excel formulas recalculate completed visits, cancellations,
no-shows and the difference from the first option. No macros are used.

**Schedule Comparison** contains distinct patient counts for the configured
windows. Custom distinct-patient counts require a report rerun: summing visits
would double count people, and this release deliberately avoids requiring modern
Excel dynamic-array functions.

```r
# Unprotected Excel exports, from the project folder:
Rscript run_report.R --mode basic --unprotected --as-of 2026-08-30

# Synthetic demonstration; never point this at hospital source folders:
Rscript examples/run_demo.R

# Decision-rule checks:
Rscript tests/run_tests.R
```

## Interpret the opening-hours comparison correctly

All options use the **same dates with clinic records** as the denominator.
Those dates are not a verified operating calendar. Opening is inclusive and
closing is exclusive, based on appointment start time.

The calculation retains observed appointment times. It does not assume that
patients move to different hours, that unoffered hours have no demand, or that
changing opening hours causes the observed completion rate. A visit starting
before closing may finish after closing: duration is not in the Basic export.

8–5 is nine clock hours; 8–4 and 9–5 are eight. More recorded activity in a longer
window does not by itself establish better staffing efficiency. Added-hour
activity, visits excluded, and patients losing all recorded completed visits are
shown to support the manager's access review.

## What changed from V3

- Retains the original overview, provider, group, cancellation, raw-data,
  chart-data, reconciliation and documentation logic.
- Adds current activity, an editable opening-hours comparison, schedule-level
  distinct patient counts, recommendations, evidence and rule documentation.
- Adds Full capacity inputs and conditional transfer scenarios.
- Supports unprotected `.xlsx` imports as well as the original Windows/Excel
  protected-workbook route. No runtime package installation or automatic updates.
- Uses explicit date-order configuration; fixes the 16:01–16:14 time-category
  error; rejects generic positional headers; gives each shared-drive run a unique
  output folder.
- Excludes flagged and unresolved records from decision calculations while
  retaining them in the original audit/detail sheets. Consequently, decision
  totals can differ from the original inclusive overview; the exclusion counts
  make that difference visible.

## Full mode

The synthetic demo creates `templates/planning_inputs_blank.xlsx` with keys from
the fictional clinics. For real clinics, generate a template from that run's
evidence; see [Full inputs](docs/FULL_INPUTS.md). Blank inputs stay blank.

Full also compares supplied opening-hour plans against a stated budget. The report
retains the planning inputs for audit. For resource transfers, this release assesses a small transfer of staffed hours between time blocks for
the same clinic, visit type and service format. It is not a complete staffing
optimizer. It subtracts the source's expected lost activity, caps destination
activity by stated demand and slots, checks protected hours and budget, and
requires access/case-mix review. Assumption ranges are not statistical confidence
intervals. Options are alternatives and must not be added together.

## Status and next gate

This is an initial, testable enhancement. Core rules and both unprotected Excel
flows are tested with synthetic records. The Windows Excel automation and actual
hospital shared-drive policies still need a workstation acceptance run. No real
patient dataset was used in development.

First acceptance test: reconcile one completed reporting period against the
existing workbook, confirm the opening/closing boundaries with the clinic
manager, and compare one proposed schedule. See [acceptance criteria](docs/ACCEPTANCE.md).

Generated operational reports remain sensitive even when MRN and Patient columns
are omitted. Provider details, dates and free-text cancellation reasons remain.
Keep hospital inputs/outputs in access-controlled folders. The repository is
public and contains only code, documentation and explicitly fictional examples.
Real patient data and operational reports must stay out of Git history.

## Development and releases

Run `Rscript tests/run_tests.R` for the core rules and
`Rscript tests/test_interface.R` for the interface adapters and server state.
The second command needs the development-only `testthat` package.

`Rscript scripts/package_release.R` creates two archives from the same source:
the hospital copy-and-run folder and the complete development source. It packages
an explicit hospital file list and Git-tracked development files, excluding input
and output data. See [development](docs/DEVELOPMENT.md).

Working project/repository name: `clinic-capacity-planner`. Hosting is a later
option; both versions currently run locally. GitHub hosts source, not an R runtime.
