# Full planning inputs

Basic always works without this workbook. Full adds inputs the original export
does not contain; missing values never mean zero.

After a Basic run in R, generate a blank template for the actual current scopes:

```r
options(clinic_capacity_no_autorun = TRUE)
source("run_report.R")
run <- cc_run(character(), getwd())
source("R/workbook_support.R")
source("R/decision_support.R")
cc_write_full_template("planning_inputs.xlsx", run$decision_result)
```

This refuses to overwrite an existing planning workbook. If the date window
changes, reconcile a new template; never reuse denominators from a different period.

## Hours Plans sheet: compare 8–4, 8–5 and 9–5 within a budget

Each row uses the exact `clinic` and `option` from Schedule Comparison, the same
`period_start` / `period_end`, and three numeric inputs: `planned_staffed_hours`,
`planned_total_cost`, and `budget_limit`. Enter totals for the whole comparison
period. `access_reviewed` and `duration_reviewed` must be TRUE after those reviews.

Costed Hours combines the observed activity captured by each window with those
planned resources. It calculates cost per captured completed visit and captured
visits per planned staffed hour, and identifies the highest visit capture among
eligible supplied options within the same budget. This ranks only supplied,
reviewed alternatives; it does not prove a universally optimal schedule.
Missing evidence for added hours, inconsistent budgets, unreviewed access/end
times, invalid inputs, and over-budget options are not eligible for ranking.

## Capacity sheet

Each row is one clinic + visit type + individual/inferred-group format + disjoint
time block. Keys come directly from Decision Evidence.

| Field | Definition |
|---|---|
| evidence_key | Exact scope key copied from the generated template |
| period_start / period_end | Exact current-analysis dates, as ISO text |
| staffed_hours | Total allocated staff hours in that scope and period, including each jointly assigned staff member's time |
| bookable_patient_slots | Patient places offered across the period; cancellations and replacement bookings do not create extra slots |
| allocated_cost | Total resource cost allocated to that scope and period, in the same currency across rows |
| protected_min_hours | Minimum hours to preserve in the source scope |
| group_confirmed | TRUE only after the inferred group structure is checked locally |

Allocate shared hours/cost once; the data cannot detect double-counted allocations
across different keys. Group participants are not separate staff sessions. Use
patient-place capacity for groups and aggregate staff hours including co-facilitators.
Incomplete/duplicate keys, zero hours, impossible slots, and period mismatches are withheld.

## Scenarios sheet

| Field | Definition |
|---|---|
| scenario | A readable option name |
| from_key / to_key | Source and destination scopes; same clinic, type and format |
| hours_to_move | Staff hours to transfer over the whole comparison period |
| additional_patient_demand | Additional suitable patient contacts that could be served in the destination during that period |
| additional_slots | Additional patient places possible in the destination |
| destination_extra_hours_limit | Maximum additional staff hours the destination can actually accommodate |
| realisation_low / realisation_high | Assumed fractions of historical destination output per staff hour, each between 0 and 1 |
| source_avoidable_cost_per_hour | Actual cost avoided per moved source hour; usually zero for unchanged salaried staffing |
| destination_incremental_cost_per_hour | Actual added cost per destination hour |
| implementation_cost | One-time incremental cost of this option |
| allowed_budget_increase | Maximum permitted incremental cost for the whole period |
| access_reviewed / case_mix_reviewed | TRUE after an actual access and comparability review |

These are totals for the analysis period, not weekly values. The model estimates:

`net additional completed visits = min(eligible demand, added slots, moved hours × destination completed/staffed hour × realisation) − moved hours × source completed/staffed hour`

`incremental cost = moved hours × (destination incremental cost/hour − source avoidable cost/hour) + implementation cost`

The low/high values are sensitivity assumptions. They are not statistical
confidence bounds and do not prove a change will cause the estimated gain.
The default pilot cap is 20% of source staff hours. Protected hours and destination
limits also apply. The source productivity is assumed constant for the hours
removed; assess patient access and scheduling feasibility separately.

Full currently screens small hours transfers. A complete schedule/budget solver,
new-hour demand forecasting, appointment-duration scheduling, case-mix adjustment,
and jointly optimising multiple alternatives are outside this first release.
