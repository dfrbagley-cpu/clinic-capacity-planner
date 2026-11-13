# Acceptance and release gate

The next decision is whether Basic's opening-hours comparison is accurate and
useful for one clinic manager. A working shared-drive reporting cycle takes
priority over hosting or further optimisation features.

1. Use a closed reporting period with the same exports as the original V3 run.
2. Reconcile original overview totals and inspect every decision exclusion.
3. Hand-check appointments at 08:00, 09:00, 16:00 and 17:00; confirm start-inclusive,
   end-exclusive rules. Check a visit at 16:10 is not labelled before 16:00.
4. Check a patient with multiple visits, a joint visit and a group session.
5. Compare 8–4, 8–5 and 9–5 over the same dates. Confirm the patient and visit
   differences against filtered source records.
6. In Excel, change a start/end input and confirm counts recalculate. Check an
   invalid window and an added hour with no observed appointments.
7. Run from the intended shared drive on the hospital Windows workstation,
   including protected imports. Confirm that two simultaneous runs get separate
   output folders and do not overwrite reports.
8. Have the manager explain which option appears preferable and which access,
   staffing or end-time question remains unresolved. Do not change clinic hours
   solely because the historical option captures more visits.

Success: counts reconcile exactly after documented exclusions; comparisons use
identical periods/days; one manager can interpret the difference without analyst
translation; no unexplained formula or import error.

Pause rollout if counts fail reconciliation, incomplete exports materially
change the preferred option, shared-drive/Excel behaviour fails, or the manager
interprets an observational comparison as guaranteed future activity.

For a pilot, monitor completed visits, distinct patients, patients unable to use
the new times, cancellations, staff hours and cost. Agree acceptable access and
quality limits locally before the pilot; no universal clinical target is embedded.

## Development verification

- 39 independent checks cover time boundaries, patient deduplication, fixed-day
  denominators, exclusions, unknown hours, missing inputs, budget/demand/slot caps,
  protected capacity, group validation and costed opening-hour comparisons.
- Both Basic and Full execute the preserved reporting pipeline against 738
  synthetic rows using the original 12-column Excel schema.
- 16 interface checks cover live distinct-patient counts, reference changes,
  invalid hours, upload handling, escaped source labels, session separation and
  recovery from failed imports. The complete Shiny UI and app entry point also
  construct successfully with Shiny 1.8.0.
- The actual Excel workflow is checked against the live interface calculations
  in both modes, including export of custom comparisons.
- Windows/Excel protected-file automation and the actual hospital shared-drive
  environment require local acceptance. No production readiness claim is made.

## Interface acceptance

Use the fictional example first. Load Basic, select each clinic, change the
reference and custom hours, and check that visits and distinct patients match
filtered source data. Try an invalid time and an unreadable replacement file;
the previous results must not appear as a successful new import. Download the
current comparison, original report and Full template; check their scopes.

Repeat in Full with the fictional planning workbook. Verify that its longer
over-budget option is not ranked as the preferred eligible option, and that
changing Custom hours does not borrow a saved plan's costs. Check keyboard use,
200% text enlargement and the intended browser/window size. Browser testing is
not part of the automated checks listed above.

## Design references

The separation of required/available capacity and observed activity is informed
by NHS England's demand-and-capacity models:
https://www.england.nhs.uk/ourwork/demand-and-capacity/service-level-demand-and-capacity-planning/models/

Access and quality measures accompany efficiency changes, consistent with IHI's
use of balancing measures:
https://www.ihi.org/library/model-for-improvement/establishing-measures

Neither source validates this tool's local screening thresholds or establishes
Ontario reporting requirements. Thresholds are configurable development defaults.

## Protected Excel acceptance on Windows (synthetic data only)

The automated smoke workflow uses unprotected workbooks. It does not test
`excel.link`, desktop Excel automation, password prompts or hospital permissions.
Run this separate check on the intended Windows workstation before local use:

1. Copy `examples/synthetic_input/visits.xlsx` into a new approved test folder.
   Never use patient exports for this check. Preserve the original fixture.
2. In desktop Excel, use **File > Info > Protect Workbook > Encrypt with Password**
   on the copy. Set a temporary test password, save and close Excel. Sheet
   protection is not file encryption. Keep the password out of scripts and logs.
3. From the project checkout, first run an unprotected baseline:
   `Rscript run_report.R --unprotected --input examples/synthetic_input --output output/acceptance-baseline --as-of 2026-08-30`.
   Run Basic with `protected = TRUE` against the encrypted-copy folder, the same
   analysis date and a different output folder. Supply the password only through
   the masked prompt. Do not add `--unprotected` to this second run.
4. Compare the resulting Schedule Comparison, Current Activity and Decision
   Evidence sheets against the baseline. Counts and exclusions must match;
   generated timestamps/output paths need not. Meridian's three saved windows
   have 112, 160 and 144 completed visits in this fixture.
5. Try an incorrect password and cancel the prompt. Each must fail clearly,
   without a usable success report or a printed password. Then retry with the
   correct password to verify recovery. Confirm source files are unchanged and
   no Excel workbook remains locked after success or failure.
6. Repeat the successful run from the approved shared-drive location; run twice
   concurrently and verify separate output folders. Repeat in Full mode with
   `examples/synthetic_planning_inputs.xlsx` and compare its baseline as well.
7. Record date, source commit/release, Windows/Excel versions, the environment
   inventory described in [ENVIRONMENT.md](ENVIRONMENT.md), results, discrepancies
   and reviewer. Keep the record locally; do not mark this checklist passed until
   it has actually been executed. Stop rollout on an unexplained mismatch.

This acceptance check establishes local technical behavior, not clinical
validation, patient-data authorization or approval of operational decisions.
