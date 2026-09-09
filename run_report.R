# Source in the R project, or run: Rscript run_report.R --mode basic
cc_run <- function(args = commandArgs(trailingOnly = TRUE), project_root = NULL,
                   config_overrides = list(), password_provider = NULL,
                   keep_analysis_data = FALSE) {
  if (is.null(project_root)) {
    file_arg <- grep("^--file=", commandArgs(), value = TRUE)
    project_root <- if (length(file_arg)) dirname(normalizePath(sub("^--file=", "", file_arg[1]))) else getwd()
  }
  project_root <- normalizePath(project_root, mustWork = TRUE)
  previous <- getwd(); on.exit(setwd(previous), add = TRUE); setwd(project_root)
  source("config.R", local = TRUE)
  if (!is.list(config_overrides) || (length(config_overrides) &&
      (is.null(names(config_overrides)) || any(!nzchar(names(config_overrides))))))
    stop("Configuration overrides must be a named list.")
  unknown_settings <- setdiff(names(config_overrides), c(names(settings), "schedules", "decision_rules"))
  if (length(unknown_settings)) stop("Unknown setting: ", unknown_settings[1])
  for (name in names(config_overrides)) settings[[name]] <- config_overrides[[name]]
  if (!is.null(password_provider) && !is.function(password_provider))
    stop("password_provider must be a function.")
  get_arg <- function(flag, default) {
    p <- match(flag, args)
    if (is.na(p)) return(default)
    if (p == length(args) || startsWith(args[p + 1L], "--")) stop("Missing value for ", flag)
    args[p + 1L]
  }
  allowed <- c("--mode", "--input", "--output", "--as-of", "--full-inputs", "--unprotected")
  unknown <- args[startsWith(args, "--") & !args %in% allowed]
  if (length(unknown)) stop("Unknown option: ", unknown[1])
  settings$mode <- tolower(get_arg("--mode", settings$mode))
  settings$input_directory <- get_arg("--input", settings$input_directory)
  settings$output_directory <- get_arg("--output", settings$output_directory)
  settings$full_inputs <- get_arg("--full-inputs", settings$full_inputs)
  settings$as_of <- as.Date(get_arg("--as-of", as.character(settings$as_of)))
  if ("--unprotected" %in% args) settings$protected <- FALSE
  if (!settings$mode %in% c("basic", "full")) stop("Mode must be basic or full.")
  if (is.na(settings$as_of)) stop("Use --as-of YYYY-MM-DD.")
  if (!settings$date_order %in% c("mdy", "dmy")) stop("date_order must be mdy or dmy.")
  dir.create(settings$output_directory, recursive = TRUE, showWarnings = FALSE)
  settings$output_directory <- normalizePath(settings$output_directory, mustWork = TRUE)
  settings$input_directory <- normalizePath(settings$input_directory, mustWork = TRUE)
  if (identical(settings$input_directory, settings$output_directory)) stop("Input and output folders must differ.")
  # Each run owns its destination. Concurrent shared-drive runs cannot overwrite one another.
  run_dir <- tempfile(paste0("run_", format(Sys.time(), "%Y%m%d_%H%M%S"), "_"), tmpdir = settings$output_directory)
  if (!dir.create(run_dir)) stop("Cannot create the report folder. Check shared-drive write access.")
  settings$output_directory <- run_dir
  source("R/decision_support.R", local = TRUE)
  source("R/workbook_support.R", local = TRUE)
  decision_config <- cc_defaults()
  decision_config$mode <- settings$mode
  decision_config$as_of <- settings$as_of
  decision_config$lookback_days <- settings$lookback_days
  if (!is.null(settings$schedules)) decision_config$schedules <- settings$schedules
  # Optional tuning uses named values; no missing Full field is ever invented.
  if (!is.null(settings$decision_rules)) for (name in names(settings$decision_rules)) {
    if (!name %in% names(decision_config)) stop("Unknown decision rule: ", name)
    decision_config[[name]] <- settings$decision_rules[[name]]
  }
  workbook_password <- NULL
  on.exit({ workbook_password <- NULL }, add = TRUE)
  cc_password <- function() {
    if (!is.null(password_provider)) return(password_provider())
    if (requireNamespace("rstudioapi", quietly = TRUE) && rstudioapi::isAvailable())
      return(rstudioapi::askForPassword("Enter the shared source-workbook password"))
    if (requireNamespace("getPass", quietly = TRUE)) return(getPass::getPass("Workbook password: "))
    stop("For protected exports, run in RStudio or ask IT to install getPass for a masked password prompt.", call. = FALSE)
  }
  full_capacity <- full_scenarios <- full_hours_plans <- NULL
  if (settings$mode == "full" && file.exists(settings$full_inputs)) {
    if (!requireNamespace("openxlsx", quietly = TRUE)) stop("openxlsx is required.")
    sheets <- openxlsx::getSheetNames(settings$full_inputs)
    if ("Capacity" %in% sheets) full_capacity <- openxlsx::read.xlsx(settings$full_inputs, sheet = "Capacity", check.names = FALSE, detectDates = TRUE)
    if ("Scenarios" %in% sheets) full_scenarios <- openxlsx::read.xlsx(settings$full_inputs, sheet = "Scenarios", check.names = FALSE, detectDates = TRUE)
    if ("Hours Plans" %in% sheets) full_hours_plans <- openxlsx::read.xlsx(settings$full_inputs, sheet = "Hours Plans", check.names = FALSE, detectDates = TRUE)
  }
  source("R/report_pipeline.R", local = TRUE, encoding = "UTF-8")
  result <- list(output_file = output_file, decision_result = decision_result)
  if (isTRUE(keep_analysis_data)) {
    result$analysis_data <- as.data.frame(visits[, cc_visit_columns])
    result$planning_inputs <- list(capacity = full_capacity, scenarios = full_scenarios,
                                  hours_plans = full_hours_plans)
  }
  invisible(result)
}

if (!isTRUE(getOption("clinic_capacity_no_autorun", FALSE))) cc_run()
