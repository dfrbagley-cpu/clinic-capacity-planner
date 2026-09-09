options(clinic_capacity_no_autorun = TRUE)
source("run_report.R")
source("R/decision_support.R")
source("R/workbook_support.R")
source("app/helpers.R")
source("app/ui.R")
source("app/server.R")
if (!requireNamespace("shiny", quietly = TRUE) || !requireNamespace("testthat", quietly = TRUE))
  stop("Interface tests need shiny and the development-only testthat package.")

checks <- 0L
check <- function(label, condition) {
  if (!isTRUE(condition)) stop("FAILED: ", label)
  checks <<- checks + 1L; cat("PASS", label, "\n")
}
fails <- function(expr) inherits(tryCatch({ force(expr); NULL }, error = identity), "error")
config <- cc_defaults(); config$as_of <- as.Date("2026-08-30")
config$min_appointments <- config$min_weeks <- 1L
visits <- cc_df(visit_record_id = 1:6, mrn = c("P1", "P2", "P1", "P3", "P3", "P4"),
  clinic_key = "C1", clinic_name = "Fictional clinic", visit_date = rep(as.Date("2026-08-24"), 6),
  appointment_time_seconds = c(8, 9, 15.5, 16, 16.5, 17) * 3600,
  visit_type_grouping_key = "T1", internal_visit_type_name = "Follow-up",
  appt_status_standardized = c("comp", "comp", "comp", "comp", "can", "no show"),
  cancellation_date = as.Date(rep(NA, 6)), possible_duplicate_appointment = FALSE,
  data_quality_issue = NA_character_, potential_group_visit = FALSE, session_id = paste0("S", 1:6))
run <- list(decision_result = cc_analyse(visits, config), analysis_data = visits,
  planning_inputs = list(capacity = NULL, scenarios = NULL, hours_plans = NULL))
check("24:00 is a valid closing boundary", cc_parse_clock("24:00", TRUE) == 86400)
check("24:00 cannot be an opening boundary", fails(cc_parse_clock("24:00")))
check("ambiguous and malformed times are rejected", fails(cc_parse_clock("8pm")) && fails(cc_parse_clock("08:60")))
check("overnight custom hours are rejected", fails(cc_app_compare(run, "17:00", "08:00")))
custom <- tail(cc_app_compare(run, "09:00", "17:00")$schedules$comparison, 1)
check("custom counts recompute distinct patients and outcomes", custom$completed_visits == 3 && custom$patients_served == 3 && custom$cancellations == 1 && custom$no_shows == 0)
changed <- cc_app_compare(run, "09:00", "16:00", "8-5")$schedules$comparison
check("changing the reference recalculates deltas", changed$option[1] == "8-5" && tail(changed$change_vs_reference, 1) == -2)
check("R and interface preset counts agree", identical(cc_app_compare(run, "09:00", "17:00")$schedules$comparison$completed_visits[1:3], run$decision_result$schedules$comparison$completed_visits))
full_run <- run; full_run$decision_result$config$mode <- "full"
full_run$decision_result$full$costed_hours <- cc_df(option = "8-4", planned_total_cost = 6400)
check("custom hours never inherit a saved plan cost", identical(cc_app_compare(full_run, "11:00", "13:00")$full$costed_hours, full_run$decision_result$full$costed_hours))

stage <- tempfile("cc_input_test_"); dir.create(stage)
source_file <- tempfile(fileext = ".xlsx"); writeLines("synthetic bytes", source_file)
cc_app_stage(cc_df(name = "../../escape.xlsx", datapath = source_file), stage)
check("upload names cannot escape the session folder", file.exists(file.path(stage, "visits_001.xlsx")) && length(list.files(stage)) == 1L)
check("non-Excel inputs are rejected", fails(cc_app_stage(cc_df(name = "data.exe", datapath = source_file), stage)))
unlink(c(stage, source_file), recursive = TRUE)
check("displayed source labels are HTML-escaped", grepl("&lt;script&gt;", as.character(cc_app_table(cc_df(clinic = "<script>"))), fixed = TRUE))

# Real Shiny server lifecycle, with a small deterministic import substitute.
# Full Excel imports are exercised separately by smoke_workflows.R.
fake_runner <- function(...) run
shiny::testServer(cc_app_server(normalizePath("."), runner = fake_runner), {
  session$setInputs(mode = "basic", demo = 0, analyse = 0)
  session$setInputs(demo = 1)
  check("a demo import creates session-local results", !is.null(current()) && current()$is_demo)
  session$setInputs(clinic = "Fictional clinic | C1", reference = "8-4", custom_start = "09:00", custom_end = "17:00")
  check("server controls reach the shared engine", tail(selected()$completed_visits, 1) == 3L)
  session$setInputs(custom_end = "16:00")
  check("editing a control updates results without reimporting", tail(selected()$completed_visits, 1) == 2L)
  session$setInputs(analyse = 1)
  check("a failed replacement import clears stale results", is.null(current()) && !is.null(error()))
})
shiny::testServer(cc_app_server(normalizePath("."), runner = fake_runner), {
  check("a new session cannot see another session's records", is.null(current()))
})
cat("\n", checks, " interface checks passed.\n", sep = "")
