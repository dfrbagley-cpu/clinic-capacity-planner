# Clinic Capacity Planner: project overview

## The operational question

A visit report describes what a clinic delivered. A manager also needs to ask
whether different operating hours would serve patients better within the
resources available. Comparing 08:00–16:00, 08:00–17:00 and 09:00–17:00 requires
more than picking the largest visit count: the windows differ in length, patients
can attend repeatedly, and missing activity does not establish missing demand.

This project extends the original Visit Analysis Tool V3 into an explainable
comparison workflow. Its primary users are clinic managers and the analysts who
support their decisions. Staff reporting remains available in the original
workbook, while the added decision layer focuses on clinic activity and hours.

## What is built

| Component | What it does |
|---|---|
| Hospital R Tool | Reads the original Excel visit exports and writes a complete reporting workbook; the folder can live on a hospital shared drive |
| Interface Version | Uses Shiny to load those exports, select a clinic, edit opening hours and recalculate visits and distinct patients |
| Basic mode | Compares observed activity using only the original export fields |
| Full mode | Adds supplied staffing, capacity, cost, demand and budget assumptions to assess costed hours and constrained staff-hour transfers |
| Shared engine | Keeps the reporting and interface calculations aligned; a UI change does not create a second set of metric definitions |

## Design decisions worth discussing

**Start with the data an analyst actually has.** Basic requires the same 12 export
columns as the original tool. The richer Full inputs are optional because a
useful initial workflow cannot depend on staffing or demand data that may not
yet be available.

**Make the first deployment fit the hospital workflow.** R runs on the analyst's
workstation, shared-drive folders hold input/output files, and managers can use
the Excel report. The interface is a second delivery route over the same engine.
An internet service is not required during ordinary operation.

**Keep the denominator comparable.** Opening-hour options use the same reporting
period and dates with clinic records. Those dates are explicitly not described
as a verified operating calendar. Visits per clock hour are not labelled as
staff productivity; staff-hour denominators require Full inputs.

**Count people directly.** Distinct patients are recomputed for every saved or
interactive window. Adding patient counts from separate time blocks would double
count people. Group attendances, inferred sessions and individual visits remain
separate measures.

**Expose exclusions and assumptions.** The decision calculations exclude flagged
or unresolved records while the original audit sheets retain them. Differences
from the inclusive overview are therefore visible and reviewable. Added hours
with no observed appointments are marked as unknown demand.

**Make recommendations conditional.** Historical visits are retained at their
recorded times. The tool does not predict rescheduling, establish causality or
verify that appointments finish before closing. A proposed staff-hour transfer
subtracts expected activity lost at the source and checks demand, slots,
protected hours, budget and access assumptions at the destination.

## Evidence and practical limits

The included demonstration uses **738 fictional records** in the original
12-column schema. Automated verification covers **39 calculation checks** and
**16 interface checks**, plus actual Basic and Full Excel imports, report exports,
and comparison of interface results with the reporting engine.

The synthetic checks establish behavior on defined cases. They do not establish
clinical benefit, savings, predictive accuracy or production readiness. This
enhanced release still needs reconciliation against an approved hospital period,
Windows/Excel protected-import testing and manager usability acceptance.

## How to review it

1. Read the example in the [README](../README.md).
2. Start the [Interface Version](INTERFACE_SETUP.md) and load the fictional example.
3. Change the clinic, reference window and custom start/end times. Inspect both
   captured visits and distinct patients, including activity excluded.
4. Load the Full example and inspect which supplied options exceed the budget.
5. Review the [shared calculations](../R/decision_support.R),
   [core checks](../tests/run_tests.R) and [acceptance criteria](ACCEPTANCE.md).

The next priority is operational validation of one completed reporting period,
followed by a small, locally reviewed clinic-hours pilot. More advanced prediction
or hosted deployment should follow evidence that it would improve the decision.
