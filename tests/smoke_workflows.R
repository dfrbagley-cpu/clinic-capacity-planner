# Slower integration check: actual Excel imports, both modes, and UI parity.
options(clinic_capacity_no_autorun = TRUE)
source("run_report.R")
source("R/decision_support.R")
source("R/workbook_support.R")
source("app/helpers.R")
root <- normalizePath(".")
stage <- tempfile("cc_workflow_smoke_"); dir.create(stage)
tryCatch({
  results <- list()
  for (mode in c("basic", "full")) {
    results[[mode]] <- cc_run(character(), root, config_overrides = list(mode = mode,
      input_directory = "examples/synthetic_input", output_directory = file.path(stage, mode),
      protected = FALSE, as_of = as.Date("2026-08-30"), lookback_days = 56L,
      full_inputs = "examples/synthetic_planning_inputs.xlsx"), keep_analysis_data = TRUE)
    result <- results[[mode]]
    stopifnot(file.exists(result$output_file), identical(names(result$analysis_data), cc_visit_columns))
    live <- cc_app_compare(result, "09:00", "17:00")
    saved <- result$decision_result$schedules$comparison
    live_presets <- live$schedules$comparison[live$schedules$comparison$option %in% saved$option, ]
    stopifnot(identical(saved$completed_visits, live_presets$completed_visits),
      identical(saved$patients_served, live_presets$patients_served))
    meridian <- saved[grepl("Meridian", saved$clinic), ]
    stopifnot(identical(meridian$completed_visits, c(112L, 160L, 144L)))
    cc_write_comparison(file.path(stage, paste0(mode, "-comparison.xlsx")), live)
    counts <- openxlsx::read.xlsx(file.path(stage, paste0(mode, "-comparison.xlsx")), sheet = "Schedule Comparison")
    stopifnot(nrow(counts) == nrow(live$schedules$comparison))
    cat("PASS actual", mode, "Excel import, report export, interface parity and custom comparison export\n")
  }
  stopifnot(nrow(results$basic$decision_result$full$costed_hours) == 0L,
    nrow(results$full$decision_result$full$costed_hours) > 0L)
  cat("PASS both modes retain independent input requirements\n")
}, finally = unlink(stage, recursive = TRUE))
