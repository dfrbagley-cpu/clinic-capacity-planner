# From the project folder: Rscript examples/run_demo.R
options(clinic_capacity_no_autorun = TRUE)
source("run_report.R")
source("examples/make_synthetic.R")
cc_make_synthetic("examples/synthetic_input/visits.xlsx")
basic <- cc_run(c("--mode", "basic", "--input", "examples/synthetic_input",
                  "--output", "output/demo_basic", "--as-of", "2026-08-30", "--unprotected"), getwd())
source("R/decision_support.R")
source("R/workbook_support.R")
template <- "templates/planning_inputs_blank.xlsx"
if (file.exists(template)) unlink(template)
cc_write_full_template(template, basic$decision_result)
e <- basic$decision_result$evidence
e <- e[e$period == "Current", ]
capacity <- cc_df(evidence_key=e$evidence_key, period_start="2026-07-06", period_end="2026-08-30",
  staffed_hours=ifelse(e$time_block=="8-9",24,32), bookable_patient_slots=pmax(e$analysed_appointments*2,1),
  allocated_cost=ifelse(e$time_block=="8-9",24,32)*65, protected_min_hours=8, group_confirmed=TRUE)
source_key <- e$evidence_key[e$clinic=="Meridian Clinic" & e$time_block=="8-9" & e$service_format=="Individual"]
target_key <- e$evidence_key[e$clinic=="Meridian Clinic" & e$time_block=="9-16" & e$service_format=="Individual"]
scenario <- cc_df(scenario="Synthetic: shift four staff hours", from_key=source_key, to_key=target_key,
  hours_to_move=4, additional_patient_demand=20, additional_slots=20, destination_extra_hours_limit=8,
  realisation_low=0.7, realisation_high=0.9, source_avoidable_cost_per_hour=0,
  destination_incremental_cost_per_hour=0, implementation_cost=0, allowed_budget_increase=0,
  access_reviewed=TRUE, case_mix_reviewed=TRUE)
s <- basic$decision_result$schedules$comparison
hours_plans <- cc_df(clinic=s$clinic,option=s$option,period_start="2026-07-06",period_end="2026-08-30",
  planned_staffed_hours=s$window_hours*s$comparison_days*2,
  planned_total_cost=s$window_hours*s$comparison_days*2*50,budget_limit=6400,
  access_reviewed=TRUE,duration_reviewed=TRUE)
openxlsx::write.xlsx(list("Hours Plans"=hours_plans, Capacity=capacity, Scenarios=scenario), "examples/synthetic_planning_inputs.xlsx", overwrite=TRUE)
full <- cc_run(c("--mode", "full", "--input", "examples/synthetic_input", "--output", "output/demo_full",
                 "--as-of", "2026-08-30", "--unprotected", "--full-inputs", "examples/synthetic_planning_inputs.xlsx"), getwd())
saveRDS(list(basic=basic, full=full), "output/demo_results.rds")
cat("BASIC_DEMO=", basic$output_file, "\nFULL_DEMO=", full$output_file, "\n", sep="")
