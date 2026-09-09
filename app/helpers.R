# Interface adapters only. Clinical activity rules live in R/decision_support.R.
cc_parse_clock <- function(value, closing = FALSE) {
  if (length(value) != 1L || is.na(value) ||
      !grepl("^([01][0-9]|2[0-3]):[0-5][0-9]$", value)) {
    if (isTRUE(closing) && identical(value, "24:00")) return(86400)
    stop("Enter a time in 24-hour HH:MM format, such as 08:00 or 17:00.", call. = FALSE)
  }
  parts <- as.numeric(strsplit(value, ":", fixed = TRUE)[[1]])
  parts[1] * 3600 + parts[2] * 60
}

cc_clock_label <- function(seconds) {
  seconds <- round(seconds)
  sprintf("%02d:%02d", as.integer(seconds %/% 3600), as.integer(seconds %% 3600 %/% 60))
}

cc_app_compare <- function(run, start, end, reference = NULL) {
  config <- run$decision_result$config
  a <- cc_parse_clock(start); b <- cc_parse_clock(end, closing = TRUE)
  if (b <= a) stop("Closing time must be later than opening time on the same day.", call. = FALSE)
  saved <- config$schedules
  if (!is.null(reference) && reference %in% saved$option)
    saved <- saved[c(which(saved$option == reference), which(saved$option != reference)), , drop = FALSE]
  custom_name <- "Custom window"
  while (custom_name %in% saved$option) custom_name <- paste0(custom_name, " *")
  config$schedules <- rbind(saved, data.frame(option = custom_name, start = a, end = b))
  planning <- run$planning_inputs
  result <- cc_analyse(run$analysis_data, config, planning$capacity, planning$scenarios, planning$hours_plans)
  # Cost inputs refer to saved schedules. A moving Custom window never inherits
  # the cost of a differently timed plan with the same label.
  result$full$costed_hours <- run$decision_result$full$costed_hours
  result$audit <- rbind(result$audit, cc_df(item = "Interface costed-hours scope",
    value = "Costed hours use the saved import-time schedules. Custom window has no costed plan; download the original-scope Full template to supply plans."))
  result
}

cc_app_stage <- function(files, directory) {
  if (is.null(files) || !nrow(files)) stop("Choose at least one Excel visit export.", call. = FALSE)
  if (any(tolower(tools::file_ext(files$name)) != "xlsx"))
    stop("Use .xlsx workbooks with the original visit-export columns.", call. = FALSE)
  dir.create(directory, recursive = TRUE, showWarnings = FALSE)
  for (i in seq_len(nrow(files))) {
    # Original filenames never become paths inside the temporary session folder.
    if (!file.copy(files$datapath[i], file.path(directory, sprintf("visits_%03d.xlsx", i)), overwrite = FALSE))
      stop("Could not read the selected workbook. Close it in Excel and try again.", call. = FALSE)
  }
  directory
}

cc_write_comparison <- function(file, result) {
  wb <- openxlsx::createWorkbook()
  for (name in names(cc_tables(result))) {
    openxlsx::addWorksheet(wb, name, gridLines = FALSE)
    data <- cc_tables(result)[[name]]
    if (ncol(data) && nrow(data)) openxlsx::writeData(wb, name, data,
      headerStyle = openxlsx::createStyle(fgFill = "#245C70", fontColour = "#FFFFFF", textDecoration = "bold", wrapText = TRUE))
    else openxlsx::writeData(wb, name, "No applicable results for this analysis.")
  }
  cc_add_hours_sandbox(wb, result)
  cc_format_decision_sheets(wb, result)
  cc_save_workbook(wb, file)
}

cc_app_table <- function(data, columns = names(data), limit = 200L) {
  if (is.null(data) || !nrow(data)) return(shiny::tags$p(class = "empty-note", "No applicable results for this selection."))
  data <- head(data[, intersect(columns, names(data)), drop = FALSE], limit)
  display <- lapply(data, function(x) {
    if (inherits(x, "Date")) return(format(x))
    if (is.numeric(x)) return(ifelse(is.na(x), "—", format(round(x, 2), big.mark = ",", trim = TRUE)))
    ifelse(is.na(x), "—", as.character(x))
  })
  shiny::tags$div(class = "table-scroll", tabindex = "0", role = "region", `aria-label` = "Results table",
    shiny::tags$table(class = "cc-table",
      shiny::tags$thead(shiny::tags$tr(lapply(names(data), function(x)
        shiny::tags$th(scope = "col", tools::toTitleCase(gsub("_", " ", x)))))),
      shiny::tags$tbody(lapply(seq_len(nrow(data)), function(i)
        shiny::tags$tr(lapply(display, function(x) shiny::tags$td(x[i])))))))
}
