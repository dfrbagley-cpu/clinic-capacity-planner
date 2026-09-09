cc_app_server <- function(project_root, runner = cc_run) {
  force(project_root); force(runner)
  function(input, output, session) {
    current <- shiny::reactiveVal(NULL)
    error <- shiny::reactiveVal(NULL)
    session_dir <- tempfile("clinic_session_"); dir.create(session_dir)
    session$onSessionEnded(function() unlink(session_dir, recursive = TRUE))

    load_exports <- function(demo = FALSE) {
      error(NULL)
      # Invalidate old results before an import, including failed imports.
      current(NULL)
      stage <- tempfile("analysis_", tmpdir = session_dir); dir.create(stage)
      mode <- if (identical(input$mode, "full")) "full" else "basic"
      tryCatch({
        if (demo) {
          visits_dir <- file.path(project_root, "examples/synthetic_input")
          planning_file <- file.path(project_root, "examples/synthetic_planning_inputs.xlsx")
          as_of <- as.Date("2026-08-30"); weeks <- 8L; protected <- FALSE; sheet <- 1L; date_order <- "mdy"
        } else {
          visits_dir <- cc_app_stage(input$visits, file.path(stage, "input"))
          as_of <- as.Date(input$as_of); weeks <- as.integer(input$weeks)
          protected <- isTRUE(input$protected); sheet <- as.integer(input$sheet); date_order <- input$date_order
          if (length(as_of) != 1L || is.na(as_of)) stop("Choose a valid reporting end date.")
          if (length(weeks) != 1L || is.na(weeks) || weeks < 1L) stop("Choose the number of weeks to analyse.")
          if (length(sheet) != 1L || is.na(sheet) || sheet < 1L) stop("Choose a valid worksheet number.")
          if (protected && .Platform$OS.type != "windows") stop("Protected import needs Windows and desktop Excel. Use an approved unprotected export on this computer.")
          if (protected && (is.null(input$password) || !nzchar(input$password))) stop("Enter the shared workbook password under Workbook options.")
          planning_file <- file.path(stage, "no_planning_inputs.xlsx")
          if (mode == "full" && !is.null(input$planning)) {
            if (tolower(tools::file_ext(input$planning$name[1])) != "xlsx") stop("Choose a .xlsx planning workbook.")
            planning_file <- input$planning$datapath[1]
          }
        }
        result <- shiny::withProgress(message = "Reading exports and calculating the report", value = 0.1, {
          runner(character(), project_root, config_overrides = list(mode = mode,
            input_directory = visits_dir, output_directory = file.path(stage, "output"),
            as_of = as_of, lookback_days = weeks * 7L, protected = protected,
            input_sheet = sheet, date_order = date_order, full_inputs = planning_file),
            password_provider = function() input$password, keep_analysis_data = TRUE)
        })
        result$is_demo <- demo
        current(result)
      }, error = function(e) {
        error(conditionMessage(e))
        unlink(stage, recursive = TRUE)
      }, finally = {
        shiny::updateTextInput(session, "password", value = "")
        # The report and minimum analysis fields suffice after import.
        unlink(file.path(stage, "input"), recursive = TRUE)
      })
    }
    shiny::observeEvent(input$analyse, load_exports(FALSE), ignoreInit = TRUE)
    shiny::observeEvent(input$demo, load_exports(TRUE), ignoreInit = TRUE)

    output$load_status <- shiny::renderUI({
      if (!is.null(error())) return(shiny::tags$div(class = "status error", role = "alert", error()))
      x <- current()
      if (is.null(x)) return(NULL)
      shiny::tags$div(class = "status success", role = "status",
        if (x$is_demo) "Fictional example loaded." else "Your exports are loaded.",
        shiny::tags$p("Changing import settings takes effect when you analyse again."))
    })
    output$period_badge <- shiny::renderUI({
      x <- current(); if (is.null(x)) return(NULL)
      cfg <- x$decision_result$config
      shiny::tags$div(class = "period-badge",
        shiny::tags$strong(if (x$is_demo) "FICTIONAL EXAMPLE" else "YOUR EXPORTS"),
        shiny::tags$span(paste(format(cfg$as_of - cfg$lookback_days + 1, "%d %b"), "–", format(cfg$as_of, "%d %b %Y"))),
        shiny::tags$span(paste(tools::toTitleCase(cfg$mode), "analysis")))
    })
    output$workspace_content <- shiny::renderUI({
      run <- current()
      if (is.null(run)) return(shiny::tags$section(class = "empty-workspace",
        shiny::tags$h2("Compare 08:00–16:00, 08:00–17:00 and 09:00–17:00"),
        shiny::tags$p("Load your Excel exports or the fictional example. Then adjust the custom opening and closing times to compare the activity each window captures."),
        shiny::tags$div(class = "empty-columns", shiny::tags$div(shiny::tags$strong("Current activity"), "Visits, patients, cancellations and no-shows."),
          shiny::tags$div(shiny::tags$strong("Alternative hours"), "Activity added, activity excluded and patients affected."))))
      scopes <- run$decision_result$schedules$scopes
      if (!nrow(scopes)) return(shiny::tags$section(class = "empty-workspace",
        shiny::tags$h2("No usable activity in this period"),
        shiny::tags$p("Check the reporting end date and the excluded records in the original report."),
        shiny::downloadButton("original_report", "Download original report")))
      schedules <- run$decision_result$config$schedules
      shiny::tagList(
        shiny::tags$div(class = "clinic-control", shiny::selectInput("clinic", "Clinic", choices = scopes$clinic, width = "100%")),
        shiny::uiOutput("metrics"),
        shiny::tabsetPanel(id = "view",
          shiny::tabPanel("Compare hours",
            shiny::tags$section(class = "panel-surface",
              shiny::tags$div(class = "section-heading", shiny::tags$div(shiny::tags$h2("Opening-hours comparison"),
                shiny::tags$p("Visit times stay where they were recorded. No new demand or rescheduling is assumed."))),
              shiny::tags$div(class = "hours-controls",
                shiny::selectInput("reference", "Compare changes against", choices = schedules$option, selected = schedules$option[1]),
                shiny::textInput("custom_start", "Custom opens · 24-hour time", value = "09:00", placeholder = "09:00"),
                shiny::textInput("custom_end", "Custom closes · 24-hour time", value = "17:00", placeholder = "17:00")),
              shiny::uiOutput("comparison"), shiny::uiOutput("activity_chart"),
              shiny::tags$p(class = "method-note", "Opening is included; closing is excluded. Appointment duration is unknown. Visits per clock hour is an activity measure, not staff productivity."),
              shiny::downloadButton("comparison_report", "Download this comparison", class = "btn-primary"))),
          shiny::tabPanel("Recommendations", shiny::tags$section(class = "panel-surface",
            shiny::tags$h2("Actions to investigate"), shiny::uiOutput("recommendations"))),
          shiny::tabPanel("Budget & staffing", shiny::tags$section(class = "panel-surface",
            shiny::tags$h2("Plans within the supplied budget"), shiny::uiOutput("full_results"),
            shiny::downloadButton("planning_template", "Download Full input template"))),
          shiny::tabPanel("Evidence & export", shiny::tags$section(class = "panel-surface",
            shiny::tags$h2("Check the evidence"), shiny::uiOutput("evidence"),
            shiny::tags$h3("Complete reporting workbook"),
            shiny::tags$p("The original report includes the saved opening windows, staff detail and data-quality checks. Custom interface changes are in Download this comparison."),
            shiny::downloadButton("original_report", "Download original report")))))
    })

    live <- shiny::reactive({
      run <- current(); shiny::req(run, input$custom_start, input$custom_end)
      tryCatch(cc_app_compare(run, input$custom_start, input$custom_end, input$reference),
        error = function(e) shiny::validate(shiny::need(FALSE, conditionMessage(e))))
    })
    selected <- shiny::reactive({
      x <- live()$schedules$comparison; shiny::req(input$clinic)
      x[x$clinic == input$clinic, , drop = FALSE]
    })
    output$metrics <- shiny::renderUI({
      run <- current(); shiny::req(run, input$clinic)
      scopes <- run$decision_result$schedules$scopes
      s <- scopes[scopes$clinic == input$clinic, , drop = FALSE]; shiny::req(nrow(s))
      times <- run$decision_result$schedules$time_data
      times <- times[times$clinic == input$clinic, , drop = FALSE]
      values <- c(s$completed, s$patients_served, sum(times$cancellations), sum(times$no_shows))
      labels <- c("Completed visits", "Distinct patients served", "Cancellations", "No-shows")
      shiny::tagList(shiny::tags$div(class = "metric-grid", lapply(seq_along(values), function(i)
        shiny::tags$div(class = "metric", shiny::tags$span(labels[i]), shiny::tags$strong(format(values[i], big.mark = ","))))),
        shiny::tags$p(class = "activity-context", sprintf("Current activity across all recorded hours · %d dates with clinic records · %d excluded records", s$days_with_records, s$excluded_records)))
    })
    output$comparison <- shiny::renderUI({
      x <- selected(); shiny::req(nrow(x))
      shown <- cc_df(Option = x$option,
        Hours = paste(cc_clock_label(x$start * 86400), cc_clock_label(x$end * 86400), sep = "–"),
        `Clock hours` = x$window_hours, `Completed visits` = x$completed_visits,
        `Distinct patients` = x$patients_served, `Change in visits` = x$change_vs_reference,
        `Visits added` = x$completed_added, `Visits excluded` = x$completed_removed,
        `Patients outside window` = x$patients_losing_all_observed_completed_visits,
        Cancellations = x$cancellations, `No-shows` = x$no_shows,
        `Visits per comparison day` = x$completed_per_comparison_day)
      shiny::tagList(cc_app_table(shown),
        shiny::tags$p(class = "field-note", "Patients outside window means people with no completed visit inside that option, from all recorded hours. Visits added and excluded are relative to the selected reference."),
        shiny::tags$div(class = "evidence-notes", lapply(which(!x$evidence_ready | grepl("unknown", x$interpretation, fixed = TRUE)), function(i)
          shiny::tags$p(class = "status caution", paste(x$option[i], "—", x$interpretation[i])))))
    })
    output$activity_chart <- shiny::renderUI({
      x <- selected(); shiny::req(nrow(x))
      largest <- max(c(1, x$completed_visits))
      shiny::tags$div(class = "activity-bars", `aria-label` = "Completed visits by opening-hours option",
        lapply(seq_len(nrow(x)), function(i) shiny::tags$div(class = "bar-row",
          shiny::tags$span(class = "bar-label", x$option[i]),
          shiny::tags$div(class = "bar-track", shiny::tags$div(class = if (i == nrow(x)) "bar-fill custom" else "bar-fill",
            style = sprintf("width: %.3f%%", 100 * x$completed_visits[i] / largest))),
          shiny::tags$strong(x$completed_visits[i]))))
    })
    output$recommendations <- shiny::renderUI({
      result <- live(); shiny::req(input$clinic)
      r <- result$recommendations
      clinic_name <- sub(" \\| .*", "", input$clinic)
      r <- r[r$clinic %in% c("All clinics", input$clinic, clinic_name), , drop = FALSE]
      if (!nrow(r)) return(shiny::tags$p("No rule-based action is indicated for this selection."))
      shiny::tagList(shiny::tags$p(class = "method-note", "Suggestions identify a review or a small pilot. Evidence labels are screening rules, not statistical confidence levels."),
        lapply(seq_len(nrow(r)), function(i) shiny::tags$article(class = "recommendation",
          shiny::tags$span(class = "evidence-label", r$evidence_strength[i]),
          shiny::tags$h3(r$recommendation[i]), shiny::tags$p(class = "field-note", r$scope[i]),
          shiny::tags$p(r$evidence[i]), shiny::tags$p(class = "method-note", r$limits[i]),
          shiny::tags$p(shiny::tags$strong("Next step: "), r$next_step[i]))))
    })
    output$full_results <- shiny::renderUI({
      run <- current(); shiny::req(run)
      result <- run$decision_result
      if (result$config$mode != "full") return(shiny::tags$p("Basic analysis uses your original visit fields. To compare staffing costs and budgets, fill in the template, select Full, choose the planning workbook and analyse again."))
      plans <- result$full$costed_hours
      if ("clinic" %in% names(plans)) plans <- plans[plans$clinic == input$clinic, , drop = FALSE]
      shiny::tagList(shiny::tags$p(class = "method-note", "These costs apply to the saved opening windows and reporting period. The moving Custom window has no costed plan."),
        cc_app_table(plans, c("option", "captured_completed_visits", "captured_patients", "planned_staffed_hours", "planned_total_cost", "budget_limit", "assessment", "explanation")),
        shiny::tags$h3("Staff-hour transfer scenarios · all supplied clinics"), cc_app_table(result$full$scenarios),
        shiny::tags$p(class = "field-note", paste(result$full$issues, collapse = " ")))
    })
    output$evidence <- shiny::renderUI({
      run <- current(); shiny::req(run)
      shiny::tagList(cc_app_table(run$decision_result$audit),
        shiny::tags$p(class = "method-note", "All options use the same dates with clinic records. These dates are not a verified operating calendar. Distinct patients are counted directly for each window; they are not summed across time blocks."))
    })
    output$original_report <- shiny::downloadHandler(
      filename = function() "Clinic-Capacity-original-report.xlsx",
      content = function(file) { shiny::req(current()); if (!file.copy(current()$output_file, file)) stop("Could not download the report.") })
    output$comparison_report <- shiny::downloadHandler(
      filename = function() "Clinic-Capacity-current-comparison.xlsx",
      content = function(file) cc_write_comparison(file, live()))
    output$planning_template <- shiny::downloadHandler(
      filename = function() "Clinic-Capacity-planning-inputs.xlsx",
      content = function(file) { shiny::req(current()); cc_write_full_template(file, current()$decision_result) })
  }
}
