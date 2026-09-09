# Excel remains the delivery format. Existing hospital R/openxlsx reports are retained.
cc_save_workbook <- function(wb, file) {
  # Some openxlsx versions emit unused drawing relationships without drawing
  # parts. Excel tolerates them; strict readers do not. Remove ONLY unused
  # missing drawing placeholders. Missing referenced parts are a hard failure.
  stage <- tempfile("cc_export_"); dir.create(stage)
  on.exit(unlink(stage, recursive = TRUE), add = TRUE)
  raw_file <- file.path(stage, "raw.xlsx")
  openxlsx::saveWorkbook(wb, raw_file, overwrite = FALSE)
  parts <- file.path(stage, "parts"); dir.create(parts)
  zip::unzip(raw_file, exdir = parts)
  rel_files <- list.files(parts, pattern = "\\.rels$", recursive = TRUE, full.names = TRUE, all.files = TRUE)
  attr <- function(x, name) sub(paste0('.*\\b', name, '="([^"]*)".*'), '\\1', x, perl = TRUE)
  for (rel_file in rel_files) {
    xml <- paste(readLines(rel_file, warn = FALSE), collapse = "")
    owner <- file.path(dirname(dirname(rel_file)), sub("\\.rels$", "", basename(rel_file)))
    owner_text <- if (file.exists(owner) && !dir.exists(owner)) paste(readLines(owner, warn = FALSE), collapse = "") else ""
    rels <- regmatches(xml, gregexpr("<Relationship\\b[^>]+/>", xml, perl = TRUE))[[1]]
    for (rel in rels) {
      if (grepl('TargetMode="External"', rel, fixed = TRUE)) next
      target <- attr(rel, "Target")
      owner_directory <- if (basename(rel_file) == ".rels") dirname(dirname(rel_file)) else dirname(owner)
      destination <- if (startsWith(target, "/")) file.path(parts, substring(target, 2)) else file.path(owner_directory, target)
      if (file.exists(destination)) next
      id <- attr(rel, "Id"); type <- attr(rel, "Type")
      used <- any(vapply(c("r:id", "r:embed", "r:link"), function(a)
        grepl(paste0(a, '="', id, '"'), owner_text, fixed = TRUE), logical(1)))
      if (!used && grepl("/(drawing|vmlDrawing)$", type)) xml <- gsub(rel, "", xml, fixed = TRUE)
      else stop("Workbook export contains a missing referenced part: ", target)
    }
    writeLines(xml, rel_file, useBytes = TRUE)
  }
  complete <- file.path(stage, "complete.xlsx")
  zip::zipr(complete, files = list.files(parts, all.files = TRUE, no.. = TRUE), root = parts)
  if (file.exists(file)) stop("Refusing to overwrite an existing report: ", file)
  if (!file.copy(complete, file, overwrite = FALSE)) stop("Could not save the report; check shared-drive access.")
  invisible(file)
}

cc_add_hours_sandbox <- function(wb, result) {
  sheet <- "Compare Hours"
  openxlsx::addWorksheet(wb, sheet, gridLines = FALSE, zoom = 85)
  if (!nrow(result$schedules$time_data)) {
    openxlsx::writeData(wb, sheet, "No usable current-period appointments. Check the analysis end date and data quality.")
    return(invisible(wb))
  }
  openxlsx::addWorksheet(wb, "Hours Data", gridLines = FALSE)
  openxlsx::writeDataTable(wb, "Hours Data", result$schedules$time_data, tableStyle = "TableStyleMedium2")
  openxlsx::freezePane(wb, "Hours Data", firstRow = TRUE)
  openxlsx::setColWidths(wb, "Hours Data", 1, 45)
  openxlsx::setColWidths(wb, "Hours Data", 2:7, 20)
  openxlsx::addStyle(wb, "Hours Data", openxlsx::createStyle(numFmt = "hh:mm:ss"),
                    rows = 2:(nrow(result$schedules$time_data) + 1), cols = 2, gridExpand = TRUE)
  title <- openxlsx::createStyle(fontSize = 22, fontColour = "#FFFFFF", fgFill = "#17374A", textDecoration = "bold")
  header <- openxlsx::createStyle(fontColour = "#FFFFFF", fgFill = "#245C70", textDecoration = "bold", wrapText = TRUE, valign = "center")
  input <- openxlsx::createStyle(fgFill = "#FFF1B8", fontColour = "#1647A5", numFmt = "hh:mm")
  body <- openxlsx::createStyle(fontName = "Calibri", fontSize = 11, valign = "center")
  note <- openxlsx::createStyle(fontColour = "#425466", fontSize = 11, wrapText = TRUE, valign = "center")
  openxlsx::mergeCells(wb, sheet, 1:10, 1:2)
  openxlsx::writeData(wb, sheet, "Compare clinic opening hours", startRow = 1)
  openxlsx::addStyle(wb, sheet, title, rows = 1:2, cols = 1:10, gridExpand = TRUE)
  openxlsx::setRowHeights(wb, sheet, 1:2, 22)
  openxlsx::writeData(wb, sheet, "Clinic", startRow = 4)
  openxlsx::mergeCells(wb, sheet, 2:6, 4)
  openxlsx::writeData(wb, sheet, result$schedules$scopes$clinic[1], startRow = 4, startCol = 2)
  openxlsx::addStyle(wb, sheet, openxlsx::createStyle(fgFill = "#FFF1B8", fontColour = "#1647A5"), rows = 4, cols = 2:6, gridExpand = TRUE)
  openxlsx::createNamedRegion(wb, "Current Activity", name = "ClinicChoices", rows = 2:(nrow(result$schedules$scopes) + 1), cols = 1, overwrite = TRUE)
  openxlsx::dataValidation(wb, sheet, cols = 2, rows = 4, type = "list", value = "ClinicChoices")
  openxlsx::writeData(wb, sheet, "Comparison days", startRow = 4, startCol = 7)
  nscopes <- nrow(result$schedules$scopes) + 1L
  criterion <- 'SUBSTITUTE(SUBSTITUTE(SUBSTITUTE($B$4,"~","~~"),"*","~*"),"?","~?")'
  openxlsx::writeFormula(wb, sheet, sprintf("SUMIF('Current Activity'!$A$2:$A$%d,%s,'Current Activity'!$B$2:$B$%d)", nscopes, criterion, nscopes), startRow = 4, startCol = 9)
  openxlsx::mergeCells(wb, sheet, 1:10, 6)
  openxlsx::writeData(wb, sheet, paste(format(result$config$as_of - result$config$lookback_days + 1), "to", format(result$config$as_of),
      "| Yellow cells are editable. All options use the same days with recorded clinic activity."), startRow = 6)
  openxlsx::addStyle(wb, sheet, note, rows = 6, cols = 1:10, gridExpand = TRUE)
  openxlsx::setRowHeights(wb, sheet, 6, 34)
  headers <- c("Option", "Start", "End", "Clock hours", "Completed visits", "Completed / day", "Cancellations", "No-shows", "Visit change", "Evidence")
  openxlsx::writeData(wb, sheet, matrix(headers, nrow = 1), startRow = 8, colNames = FALSE)
  openxlsx::addStyle(wb, sheet, header, rows = 8, cols = 1:10, gridExpand = TRUE)
  openxlsx::setRowHeights(wb, sheet, 8, 45)
  n <- nrow(result$schedules$time_data) + 1L
  presets <- result$config$schedules
  inputs <- rbind(presets, data.frame(option = "Custom", start = 8 * 3600, end = 17 * 3600))
  for (i in seq_len(nrow(inputs))) {
    r <- 8L + i
    openxlsx::writeData(wb, sheet, inputs$option[i], startRow = r, startCol = 1)
    openxlsx::writeData(wb, sheet, matrix(c(inputs$start[i], inputs$end[i]) / 86400, nrow = 1), startRow = r, startCol = 2, colNames = FALSE)
    valid <- sprintf('AND(ISNUMBER(B%d),ISNUMBER(C%d),B%d>=0,C%d<=1,B%d<C%d)', r,r,r,r,r,r)
    openxlsx::writeFormula(wb, sheet, sprintf('IF(%s,24*(C%d-B%d),"")',valid,r,r), startRow = r, startCol = 4)
    sums <- function(col) sprintf("SUMIFS('Hours Data'!$%s$2:$%s$%d,'Hours Data'!$A$2:$A$%d,%s,'Hours Data'!$B$2:$B$%d,\">=\"&B%d,'Hours Data'!$B$2:$B$%d,\"<\"&C%d)", col,col,n,n,criterion,n,r,n,r)
    for (pair in list(c("C",5), c("D",7), c("E",8)))
      openxlsx::writeFormula(wb, sheet, sprintf('IF(%s,%s,"")',valid,sums(pair[1])), startRow = r, startCol = as.integer(pair[2]))
    openxlsx::writeFormula(wb, sheet, sprintf('IF(AND(%s,$I$4>0),E%d/$I$4,"")',valid,r), startRow = r, startCol = 6)
    openxlsx::writeFormula(wb, sheet, sprintf('IF(AND(ISNUMBER(E%d),ISNUMBER($E$9)),E%d-$E$9,"")',r,r), startRow = r, startCol = 9)
    openxlsx::writeFormula(wb, sheet, sprintf('IF(NOT(%s),"Check times",IF(%s=0,"No observed activity","Observed pattern"))',valid,sums("F")), startRow = r, startCol = 10)
    openxlsx::addStyle(wb, sheet, body, rows = r, cols = 1:10, gridExpand = TRUE)
    openxlsx::addStyle(wb, sheet, input, rows = r, cols = 2:3, gridExpand = TRUE, stack = TRUE)
    openxlsx::addStyle(wb, sheet, openxlsx::createStyle(numFmt = "0.00"), rows = r, cols = c(4,6), gridExpand = TRUE, stack = TRUE)
    openxlsx::setRowHeights(wb, sheet, r, 32)
    openxlsx::dataValidation(wb, sheet, cols = 2:3, rows = r, type = "decimal", operator = "between", value = c(0,1), allowBlank = FALSE)
  }
  notes <- c(
    "What this estimates: activity captured if recorded appointments happened at the same times again. It is a historical scenario, not a forecast of new bookings or a claim that hours caused attendance.",
    "Patient counts: Schedule Comparison contains distinct patients for the saved time windows. For different windows, update settings$schedules and rerun; visits cannot be added to infer unique patients.",
    "Timing: starts are included at opening and excluded at closing. Appointment end times are absent, so the model cannot confirm that visits would finish before closing.",
    "Evidence: an empty added hour is unknown demand, not proof of zero demand. Days with no records are not known operating days. Check Schedule Comparison for coverage and excluded records.",
    "Cost: 8-5 is nine clock hours; 8-4 and 9-5 are eight. Basic does not know staff hours or cost. Full adds those inputs. Inspect individual visits and group attendance separately before comparing workloads.")
  start_note <- 10L + nrow(inputs)
  for (i in seq_along(notes)) {
    r <- start_note + (i - 1L) * 2L
    openxlsx::mergeCells(wb, sheet, 1:10, r:(r+1L))
    openxlsx::writeData(wb, sheet, notes[i], startRow = r)
    openxlsx::addStyle(wb, sheet, note, rows = r:(r+1L), cols = 1:10, gridExpand = TRUE)
    openxlsx::setRowHeights(wb, sheet, r:(r+1L), 21)
  }
  openxlsx::setColWidths(wb, sheet, 1:10, c(17,10,10,12,16,16,15,13,17,27))
  openxlsx::freezePane(wb, sheet, firstActiveRow = 9)
  openxlsx::pageSetup(wb, sheet, orientation = "landscape", fitToWidth = 1, fitToHeight = 1)
  openxlsx::worksheetOrder(wb) <- c(match(sheet, names(wb)), setdiff(seq_along(names(wb)), match(sheet, names(wb))))
  openxlsx::activeSheet(wb) <- 1L
  # Ensure Excel recalculates native SUMIFS formulas on opening, without macros.
  wb$workbook$calcPr <- '<calcPr calcId="191029" fullCalcOnLoad="1" forceFullCalc="1" calcMode="auto"/>'
  invisible(wb)
}

cc_format_decision_sheets <- function(wb, result) {
  for (sheet in names(cc_tables(result))) {
    data <- cc_tables(result)[[sheet]]
    if (!ncol(data)) next
    headers <- tools::toTitleCase(gsub("_", " ", names(data)))
    openxlsx::writeData(wb, sheet, matrix(headers, nrow = 1), colNames = FALSE)
    openxlsx::setRowHeights(wb, sheet, 1, 72)
    openxlsx::setColWidths(wb, sheet, seq_len(ncol(data)), 21)
    text_cols <- which(vapply(data, is.character, logical(1)))
    if (length(text_cols)) openxlsx::setColWidths(wb, sheet, text_cols, 34)
    long_cols <- which(names(data) %in% c("recommendation", "evidence", "limits", "next_step", "interpretation", "explanation", "value"))
    if (length(long_cols)) openxlsx::setColWidths(wb, sheet, long_cols, 62)
    if (nrow(data)) {
      openxlsx::setRowHeights(wb, sheet, 2:(nrow(data)+1), if (length(long_cols)) 88 else 40)
      openxlsx::addStyle(wb, sheet, openxlsx::createStyle(wrapText = TRUE, valign = "top"),
        2:(nrow(data)+1), seq_len(ncol(data)), gridExpand = TRUE, stack = TRUE)
      openxlsx::addFilter(wb, sheet, rows = 1, cols = seq_len(ncol(data)))
      percent_cols <- which(grepl("rate$|fraction$|share_of", names(data)))
      if (length(percent_cols)) openxlsx::addStyle(wb, sheet, openxlsx::createStyle(numFmt = "0.0%"), 2:(nrow(data)+1), percent_cols, gridExpand = TRUE, stack = TRUE)
      decimal_cols <- which(grepl("per_|cost|hours", names(data)) & vapply(data, is.numeric, logical(1)))
      if (length(decimal_cols)) openxlsx::addStyle(wb, sheet, openxlsx::createStyle(numFmt = "0.00"), 2:(nrow(data)+1), decimal_cols, gridExpand = TRUE, stack = TRUE)
      range_cols <- which(grepl("_low$|_high$", names(data)))
      if (length(range_cols)) openxlsx::addStyle(wb, sheet, openxlsx::createStyle(numFmt = "0.0"), 2:(nrow(data)+1), range_cols, gridExpand = TRUE, stack = TRUE)
      if (sheet == "Schedule Comparison") openxlsx::addStyle(wb, sheet, openxlsx::createStyle(numFmt = "hh:mm"), 2:(nrow(data)+1), c(3,4), gridExpand = TRUE, stack = TRUE)
    }
    openxlsx::freezePane(wb, sheet, firstRow = TRUE, firstCol = TRUE)
  }
}

cc_write_full_template <- function(path, result) {
  e <- result$evidence
  if (!nrow(e)) stop("No current scopes to populate the Full template.")
  e <- e[e$period == "Current", , drop = FALSE]
  capacity <- cc_df(evidence_key = e$evidence_key,
    period_start = format(result$config$as_of - result$config$lookback_days + 1),
    period_end = format(result$config$as_of), staffed_hours = NA_real_, bookable_patient_slots = NA_real_,
    allocated_cost = NA_real_, protected_min_hours = NA_real_, group_confirmed = FALSE,
    clinic = e$clinic, visit_type = e$visit_type, service_format = e$service_format, time_block = e$time_block)
  scenario <- cc_df(scenario = "Replace with a local scenario", from_key = "", to_key = "",
    hours_to_move = NA_real_, additional_patient_demand = NA_real_, additional_slots = NA_real_,
    destination_extra_hours_limit = NA_real_, realisation_low = NA_real_, realisation_high = NA_real_,
    source_avoidable_cost_per_hour = NA_real_, destination_incremental_cost_per_hour = NA_real_,
    implementation_cost = NA_real_, allowed_budget_increase = NA_real_, access_reviewed = FALSE, case_mix_reviewed = FALSE)
  comparison <- result$schedules$comparison
  hours_plans <- cc_df(clinic=comparison$clinic, option=comparison$option,
    period_start=format(result$config$as_of-result$config$lookback_days+1), period_end=format(result$config$as_of),
    planned_staffed_hours=NA_real_,planned_total_cost=NA_real_,budget_limit=NA_real_,access_reviewed=FALSE,duration_reviewed=FALSE)
  template_data <- list("Hours Plans" = hours_plans, Capacity = capacity, Scenarios = scenario,
    Instructions = cc_df(note = c("Use actual staffed hours and costs for exactly the stated period; include all jointly assigned staff time once.",
      "Bookable slots mean patient places offered during the period, not repeated bookings of the same slot. Group slots are participant places.",
      "The template is incomplete intentionally. Blank values are never replaced with zero or invented estimates.",
      "Scenario hours, demand, room limits and costs are totals over the same period, not weekly values.",
      "Allocated cost per visit is descriptive. Cash changes use explicit marginal costs, never assumed salary savings.",
      "Realisation low/high are assumptions from 0 to 1, not statistical confidence bounds.",
      "TRUE review flags require a real access and case-mix review. Compare one scenario at a time.",
      "Do not double count overlapping staff hours or use the same budget allocation in multiple capacity rows.")))
  wb <- openxlsx::createWorkbook()
  for (sheet in names(template_data)) {
    openxlsx::addWorksheet(wb, sheet, gridLines = FALSE)
    openxlsx::writeDataTable(wb, sheet, template_data[[sheet]], tableStyle = "TableStyleMedium2")
    openxlsx::setColWidths(wb, sheet, seq_len(ncol(template_data[[sheet]])), if (sheet == "Instructions") 100 else 26)
    openxlsx::freezePane(wb, sheet, firstRow = TRUE)
  }
  cc_save_workbook(wb, path)
}
