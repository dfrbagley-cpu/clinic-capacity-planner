# Clinic Capacity Planner: deterministic decisions from the existing visit data.
# No network calls, model service, or patient-level recommendations.
cc_visit_columns <- c("visit_record_id", "mrn", "clinic_key", "clinic_name", "visit_date",
  "appointment_time_seconds", "visit_type_grouping_key", "internal_visit_type_name",
  "appt_status_standardized", "cancellation_date", "possible_duplicate_appointment",
  "data_quality_issue", "potential_group_visit", "session_id")

cc_defaults <- function() list(
  mode = "basic", as_of = Sys.Date() - 1, lookback_days = 56L,
  min_appointments = 30L, min_weeks = 4L, max_excluded_fraction = 0.05,
  no_show_threshold = 0.15, cancellation_threshold = 0.25,
  meaningful_rate_change = 0.10, advance_notice_days = 2L,
  min_advance_cancellations = 5L, max_hours_move_fraction = 0.20,
  schedules = data.frame(option = c("8-4", "8-5", "9-5"),
                         start = c(8, 8, 9) * 3600, end = c(16, 17, 17) * 3600))

cc_ratio <- function(a, b) ifelse(b > 0, a / b, NA_real_)
cc_pct <- function(x) ifelse(is.na(x), "unavailable", sprintf("%.1f%%", 100 * x))
cc_df <- function(...) data.frame(..., stringsAsFactors = FALSE, check.names = FALSE)
cc_bind <- function(rows) if (length(rows)) do.call(rbind, rows) else data.frame()
cc_week <- function(x) as.Date(x) - (as.POSIXlt(as.Date(x))$wday + 6L) %% 7L
cc_key <- function(...) {
  # Length prefixes avoid collisions when source labels contain separators.
  parts <- list(...)
  do.call(paste0, lapply(parts, function(x) paste0(nchar(as.character(x)), ":", x)))
}

cc_prepare <- function(visits, config) {
  required <- cc_visit_columns
  missing <- setdiff(required, names(visits))
  if (length(missing)) stop("Decision support is missing: ", paste(missing, collapse = ", "))
  if (anyDuplicated(visits$visit_record_id)) stop("Visit Record IDs must be unique before provider attribution.")
  if (length(config$as_of) != 1L || is.na(config$as_of)) stop("Use one valid analysis end date.")
  if (!is.numeric(config$lookback_days) || config$lookback_days < 7 || config$lookback_days %% 7 != 0)
    stop("lookback_days must be a positive whole number of weeks.")
  if (!config$mode %in% c("basic", "full")) stop("Mode must be basic or full.")
  for (name in c("min_appointments", "min_weeks", "min_advance_cancellations", "advance_notice_days")) {
    value <- config[[name]]
    if (length(value) != 1L || !is.numeric(value) || !is.finite(value) || value < 1 || value %% 1 != 0)
      stop(name, " must be a positive integer.")
  }
  for (name in c("max_excluded_fraction", "no_show_threshold", "cancellation_threshold", "meaningful_rate_change", "max_hours_move_fraction")) {
    value <- config[[name]]
    if (length(value) != 1L || !is.numeric(value) || !is.finite(value) || value < 0 || value > 1)
      stop(name, " must be between 0 and 1.")
  }
  schedules <- config$schedules
  if (!all(c("option", "start", "end") %in% names(schedules)) || !nrow(schedules) ||
      anyNA(schedules) || anyDuplicated(schedules$option) ||
      any(schedules$start < 0 | schedules$end > 86400 | schedules$start >= schedules$end))
    stop("Schedules need unique names and valid same-day start/end times.")
  d <- as.data.frame(visits)
  d$visit_date <- as.Date(d$visit_date)
  d$week <- cc_week(d$visit_date)
  d$period <- ifelse(is.na(d$visit_date), "Unknown date",
    ifelse(d$visit_date > config$as_of, "After analysis end",
    ifelse(d$visit_date > config$as_of - config$lookback_days, "Current",
    ifelse(d$visit_date > config$as_of - 2 * config$lookback_days, "Previous", "Older"))))
  status <- d$appt_status_standardized
  d$outcome <- ifelse(status %in% c("comp", "complete", "completed"), "Completed",
    ifelse(status %in% c("no show", "no-show", "noshow"), "No-show",
    ifelse(status %in% c("can", "cancelled", "canceled") | grepl("^cancel", status), "Cancelled", "Unresolved")))
  invalid <- is.na(d$visit_date) | is.na(d$appointment_time_seconds) |
    d$appointment_time_seconds < 0 | d$appointment_time_seconds >= 86400 |
    is.na(d$clinic_key) | is.na(d$visit_type_grouping_key) |
    is.na(d$mrn) | !nzchar(trimws(d$mrn))
  d$excluded <- invalid | is.na(d$possible_duplicate_appointment) |
    d$possible_duplicate_appointment | (!is.na(d$data_quality_issue) & nzchar(d$data_quality_issue)) |
    d$outcome == "Unresolved"
  d$excluded[is.na(d$excluded)] <- TRUE
  d$service_format <- ifelse(d$potential_group_visit, "Inferred group", "Individual")
  d$time_block <- as.character(cut(d$appointment_time_seconds,
    breaks = c(0, 8, 9, 16, 16.5, 24) * 3600, right = FALSE,
    labels = c("Before 8", "8-9", "9-16", "16-16:30", "16:30 onward")))
  d$advance_cancel <- d$outcome == "Cancelled" & !is.na(d$cancellation_date) &
    as.numeric(d$visit_date - as.Date(d$cancellation_date)) >= config$advance_notice_days
  d$advance_cancel[is.na(d$advance_cancel)] <- FALSE
  d
}

cc_evidence <- function(d, config) {
  d <- d[d$period %in% c("Current", "Previous") & !is.na(d$time_block), , drop = FALSE]
  if (!nrow(d)) return(data.frame())
  key <- with(d, cc_key(period, clinic_key, visit_type_grouping_key, service_format, time_block))
  rows <- lapply(split(d, key), function(x) {
    clean <- x[!x$excluded, , drop = FALSE]
    n <- nrow(clean); completed <- sum(clean$outcome == "Completed")
    excluded_fraction <- mean(x$excluded)
    weeks <- length(unique(clean$week))
    cc_df(period = x$period[1], clinic_key = x$clinic_key[1], clinic = x$clinic_name[1],
      visit_type_key = x$visit_type_grouping_key[1], visit_type = x$internal_visit_type_name[1],
      service_format = x$service_format[1], time_block = x$time_block[1],
      source_records = nrow(x), excluded_records = sum(x$excluded),
      analysed_appointments = n, completed = completed,
      patients_served = length(unique(clean$mrn[clean$outcome == "Completed"])),
      cancellations = sum(clean$outcome == "Cancelled"), no_shows = sum(clean$outcome == "No-show"),
      advance_cancellations = sum(clean$advance_cancel),
      weeks_observed = weeks, completion_rate = cc_ratio(completed, n),
      cancellation_rate = cc_ratio(sum(clean$outcome == "Cancelled"), n),
      no_show_rate = cc_ratio(sum(clean$outcome == "No-show"), n),
      excluded_fraction = excluded_fraction,
      evidence_ready = n >= config$min_appointments && weeks >= config$min_weeks &&
        excluded_fraction <= config$max_excluded_fraction)
  })
  result <- cc_bind(rows)
  result$evidence_key <- with(result, cc_key(clinic_key, visit_type_key, service_format, time_block))
  result
}

cc_schedule_analysis <- function(d, config) {
  current <- d[d$period == "Current", , drop = FALSE]
  clinics <- unique(current[, c("clinic_key", "clinic_name"), drop = FALSE])
  result <- minute_rows <- scopes <- list()
  for (i in seq_len(nrow(clinics))) {
    raw <- current[current$clinic_key == clinics$clinic_key[i], , drop = FALSE]
    x <- raw[!raw$excluded, , drop = FALSE]
    days <- length(unique(raw$visit_date[!is.na(raw$visit_date)]))
    if (!nrow(x)) next
    scope <- paste0(clinics$clinic_name[i], " | ", clinics$clinic_key[i])
    full_patients <- unique(x$mrn[x$outcome == "Completed"])
    scopes[[length(scopes) + 1L]] <- cc_df(clinic = scope, days_with_records = days,
      analysed_appointments = nrow(x), completed = sum(x$outcome == "Completed"),
      patients_served = length(full_patients), excluded_records = sum(raw$excluded),
      weeks_observed = length(unique(x$week)), source_records = nrow(raw))
    # Exact seconds are retained: no rounding across opening/closing boundaries.
    for (time in sort(unique(x$appointment_time_seconds))) {
      at <- x[x$appointment_time_seconds == time, , drop = FALSE]
      minute_rows[[length(minute_rows) + 1L]] <- cc_df(clinic = scope,
        start_time = time / 86400, completed = sum(at$outcome == "Completed"),
        cancellations = sum(at$outcome == "Cancelled"), no_shows = sum(at$outcome == "No-show"),
        appointments = nrow(at), days_with_records = length(unique(at$visit_date)))
    }
    reference <- x$appointment_time_seconds >= config$schedules$start[1] &
      x$appointment_time_seconds < config$schedules$end[1]
    for (j in seq_len(nrow(config$schedules))) {
      scenario <- config$schedules[j, ]
      inside <- x$appointment_time_seconds >= scenario$start & x$appointment_time_seconds < scenario$end
      done <- x$outcome == "Completed"
      served <- unique(x$mrn[inside & done])
      added <- inside & !reference; removed <- reference & !inside
      added_interval <- scenario$start < config$schedules$start[1] || scenario$end > config$schedules$end[1]
      ready <- nrow(x) >= config$min_appointments && length(unique(x$week)) >= config$min_weeks &&
        mean(raw$excluded) <= config$max_excluded_fraction
      note <- if (!ready) "Limited or incomplete evidence; inspect before using this comparison."
        else if (added_interval && !any(added)) "Added hours have no observed appointments; future activity there is unknown."
        else "Observed activity captured if visit times repeat; no rescheduling or new demand assumed."
      result[[length(result) + 1L]] <- cc_df(clinic = scope, option = scenario$option,
        start = scenario$start / 86400, end = scenario$end / 86400,
        window_hours = (scenario$end - scenario$start) / 3600,
        comparison_days = days, completed_visits = sum(inside & done),
        completed_individual_visits = sum(inside & done & !x$potential_group_visit),
        group_attendances = sum(inside & done & x$potential_group_visit),
        inferred_group_sessions = length(unique(x$session_id[inside & done & x$potential_group_visit & !is.na(x$session_id)])),
        patients_served = length(served), cancellations = sum(inside & x$outcome == "Cancelled"),
        no_shows = sum(inside & x$outcome == "No-show"),
        completed_per_comparison_day = cc_ratio(sum(inside & done), days),
        activity_per_window_hour = cc_ratio(sum(inside & done), days * (scenario$end - scenario$start) / 3600),
        change_vs_reference = sum(inside & done) - sum(reference & done),
        completed_added = sum(added & done), completed_removed = sum(removed & done),
        patients_losing_all_observed_completed_visits = length(setdiff(full_patients, served)),
        appointments_in_added_hours = sum(added), added_hours_observed_days = length(unique(x$visit_date[added])),
        excluded_records = sum(raw$excluded), evidence_ready = ready, interpretation = note)
    }
  }
  list(comparison = cc_bind(result), time_data = cc_bind(minute_rows), scopes = cc_bind(scopes))
}

cc_recommendations <- function(evidence, schedules, d, config) {
  rows <- list()
  add <- function(rule, clinic, scope, action, evidence_text, limit, next_step, grade = "Exploratory", priority = 2L) {
    rows[[length(rows) + 1L]] <<- cc_df(priority = priority, rule = rule, clinic = clinic,
      scope = scope, recommendation = action, evidence = evidence_text,
      evidence_strength = grade, limits = limit, next_step = next_step)
  }
  if (any(d$excluded & d$period %in% c("Current", "Unknown date")))
    add("DQ01", "All clinics", "Current period", "Resolve excluded records before interpreting small differences.",
        paste(sum(d$excluded & d$period %in% c("Current", "Unknown date")), "records excluded or undated."),
        "Original overview retains flagged rows; decision sheets exclude them. Unknown dates cannot be assigned to a period.",
        "Review Data Quality Detail and unresolved statuses, correct the export and rerun.", "Data review", 1L)
  current <- if (nrow(evidence)) evidence[evidence$period == "Current", , drop = FALSE] else evidence
  for (i in seq_len(nrow(current))) {
    e <- current[i, ]; scope <- paste(e$visit_type, e$service_format, e$time_block, sep = " / ")
    if (!e$evidence_ready) {
      add("EV01", e$clinic, scope, "Collect or verify more history before changing this service.",
          sprintf("%d analysed appointments, %d weeks; %.1f%% excluded.", e$analysed_appointments,
            e$weeks_observed, 100 * e$excluded_fraction),
          "The minimum evidence rule is a configurable screening threshold, not a statistical test.",
          "Inspect exclusions and compare again after enough comparable weeks.", "Limited")
      next
    }
    if (e$service_format == "Inferred group") {
      add("GR01", e$clinic, scope, "Confirm the inferred group structure before acting on group performance.",
          paste(e$completed, "completed patient attendances."),
          "Shared start times can represent group care, parallel activity, or scheduling conventions.",
          "Confirm group sessions with the service manager; retain attendance and session counts separately.")
      next
    }
    if (e$no_show_rate >= config$no_show_threshold)
      add("NS01", e$clinic, scope, "Test a targeted attendance-support change.",
          sprintf("%d no-shows / %d analysed appointments (%s), across %d weeks.",
            e$no_shows, e$analysed_appointments, cc_pct(e$no_show_rate), e$weeks_observed),
          "Records do not establish whether timing, transport, reminders, or other barriers caused non-attendance.",
          "Ask patients/front-line staff about barriers; test one change and track completed visits and access.")
    if (e$cancellation_rate >= config$cancellation_threshold && e$advance_cancellations >= config$min_advance_cancellations)
      add("CA01", e$clinic, scope, "Check whether earlier cancellations could support a refill-list pilot.",
          sprintf("%d cancellations; %d recorded at least %d calendar days before the visit.",
            e$cancellations, e$advance_cancellations, config$advance_notice_days),
          "Cancellation counts are not empty slots. The same slot may already have been rebooked; reasons are not causal evidence.",
          "Check a sample of cancelled slots, confirm unused capacity and eligible patients, then test the refill workflow.")
    prior <- evidence[evidence$period == "Previous" & evidence$evidence_key == e$evidence_key, , drop = FALSE]
    if (nrow(prior) == 1L && prior$evidence_ready &&
        prior$completion_rate - e$completion_rate >= config$meaningful_rate_change)
      add("TR01", e$clinic, scope, "Investigate the fall in appointment completion.",
          sprintf("Completion changed from %s (%d appointments) to %s (%d appointments).",
            cc_pct(prior$completion_rate), prior$analysed_appointments,
            cc_pct(e$completion_rate), e$analysed_appointments),
          "Adjacent equal-length periods; changes in patient mix, staffing and seasonality remain unadjusted.",
          "Review weekly evidence with clinic staff and check whether coding or the service changed.")
  }
  s <- schedules$comparison
  if (nrow(s)) for (clinic in unique(s$clinic)) {
    x <- s[s$clinic == clinic, , drop = FALSE]; base <- x[x$option == config$schedules$option[1], ]
    equal <- x[x$window_hours == base$window_hours & x$option != base$option & x$evidence_ready, , drop = FALSE]
    if (nrow(equal) && base$evidence_ready) {
      best <- equal[which.max(equal$completed_visits), ]
      if (best$completed_visits > base$completed_visits && best$patients_served >= base$patients_served)
        add("HR01", clinic, "Opening hours", paste("Consider a small", best$option, "hours pilot."),
            sprintf("Same clock span: %s captures %d completed visits and %d patients; %s captures %d and %d.",
              best$option, best$completed_visits, best$patients_served, base$option, base$completed_visits, base$patients_served),
            sprintf("Activity capture, not a causal forecast or staffing efficiency score. %d reference-window completed visits would fall outside the alternative.", best$completed_removed),
            "Check access for affected patients and end-of-visit times; test without reducing care commitments.")
    }
    longer <- x[x$window_hours > base$window_hours & x$change_vs_reference > 0 & x$evidence_ready, , drop = FALSE]
    if (nrow(longer)) {
      best <- longer[which.max(longer$change_vs_reference), ]
      add("HR02", clinic, "Opening hours", paste("Assess whether the extra hour in", best$option, "is worth funding."),
          sprintf("It captures %d additional completed visits across the same %d comparison days versus %s.",
            best$change_vs_reference, best$comparison_days, base$option),
          "More clock hours are not a free gain. Available staff, unmet demand, appointment duration and cost are unknown in Basic.",
          "Use Full for a costed scenario, or review staffing and access locally before approving an extension.")
    }
  }
  if (!length(rows)) add("MON01", "All clinics", "Current period", "Continue monitoring; no configured opportunity rule fired.",
      "See Current Activity and Schedule Comparison.", "No flag does not prove that the current schedule is optimal.",
      "Review the opening-hours comparisons and gather staff/patient feedback.")
  out <- cc_bind(rows); out[order(out$priority, out$clinic, out$rule), , drop = FALSE]
}

cc_full <- function(evidence, capacity, scenarios, config) {
  issues <- character(); efficiencies <- results <- list()
  issue <- function(text) issues <<- c(issues, text)
  if (is.null(capacity) || !nrow(capacity)) return(list(efficiency = data.frame(), scenarios = data.frame(),
    issues = "Full inputs were not supplied. Basic remains available; no cost or staffing recommendations were inferred."))
  required <- c("evidence_key", "period_start", "period_end", "staffed_hours", "bookable_patient_slots",
                "allocated_cost", "protected_min_hours", "group_confirmed")
  if (length(setdiff(required, names(capacity)))) return(list(efficiency = data.frame(), scenarios = data.frame(),
    issues = paste("Capacity inputs are missing:", paste(setdiff(required, names(capacity)), collapse = ", "))))
  c <- as.data.frame(capacity)
  duplicate_keys <- c$evidence_key[duplicated(c$evidence_key) | duplicated(c$evidence_key, fromLast = TRUE)]
  for (i in seq_len(nrow(c))) {
    row <- c[i, ]; key <- as.character(row$evidence_key)
    e <- evidence[evidence$period == "Current" & evidence$evidence_key == key, , drop = FALSE]
    numeric_cols <- c("staffed_hours", "bookable_patient_slots", "allocated_cost", "protected_min_hours")
    numbers <- suppressWarnings(as.numeric(unlist(row[numeric_cols], use.names = FALSE)))
    dates <- tryCatch(as.Date(c(as.character(row$period_start), as.character(row$period_end))), error = function(e) as.Date(c(NA, NA)))
    valid <- nrow(e) == 1L && !key %in% duplicate_keys && !anyNA(numbers) && all(is.finite(numbers)) &&
      all(numbers >= 0) && numbers[1] > 0 && numbers[2] > 0 && numbers[4] <= numbers[1] &&
      !anyNA(dates) && dates[1] == config$as_of - config$lookback_days + 1 && dates[2] == config$as_of
    if (!valid) { issue(paste("Capacity row", i, "has invalid, duplicated, unmatched or wrong-period inputs; it was withheld.")); next }
    row[numeric_cols] <- as.list(numbers)
    if (e$completed > row$bookable_patient_slots) { issue(paste("Capacity row", i, "has fewer offered slots than completed visits; it was withheld.")); next }
    efficiencies[[length(efficiencies) + 1L]] <- cbind(e, row[, setdiff(required, "evidence_key"), drop = FALSE],
      cc_df(completed_per_staffed_hour = e$completed / row$staffed_hours,
            allocated_cost_per_completed_visit = cc_ratio(row$allocated_cost, e$completed),
            completed_share_of_offered_slots = e$completed / row$bookable_patient_slots,
            full_ready = e$evidence_ready && (e$service_format == "Individual" || isTRUE(as.logical(row$group_confirmed)))))
  }
  eff <- cc_bind(efficiencies)
  if (is.null(scenarios) || !nrow(scenarios)) return(list(efficiency = eff, scenarios = data.frame(), issues = c(issues, "No Full scenarios supplied.")))
  req <- c("scenario", "from_key", "to_key", "hours_to_move", "additional_patient_demand", "additional_slots",
    "destination_extra_hours_limit", "realisation_low", "realisation_high", "source_avoidable_cost_per_hour",
    "destination_incremental_cost_per_hour", "implementation_cost", "allowed_budget_increase", "access_reviewed", "case_mix_reviewed")
  if (length(setdiff(req, names(scenarios)))) return(list(efficiency = eff, scenarios = data.frame(),
    issues = c(issues, paste("Scenario inputs missing:", paste(setdiff(req, names(scenarios)), collapse = ", ")))))
  for (i in seq_len(nrow(scenarios))) {
    z <- scenarios[i, ]; reason <- character(); low <- high <- cost <- NA_real_
    a <- if (nrow(eff)) eff[eff$evidence_key == z$from_key, , drop = FALSE] else eff
    b <- if (nrow(eff)) eff[eff$evidence_key == z$to_key, , drop = FALSE] else eff
    nums <- c("hours_to_move", "additional_patient_demand", "additional_slots", "destination_extra_hours_limit",
              "realisation_low", "realisation_high", "source_avoidable_cost_per_hour", "destination_incremental_cost_per_hour",
              "implementation_cost", "allowed_budget_increase")
    values <- suppressWarnings(as.numeric(unlist(z[nums], use.names = FALSE)))
    if (anyNA(values) || any(!is.finite(values)) || any(values < 0)) reason <- c(reason, "Missing or invalid numeric assumptions.")
    else {
      z[nums] <- as.list(values)
      if (z$hours_to_move <= 0 || z$realisation_low > z$realisation_high || z$realisation_high > 1)
        reason <- c(reason, "Invalid hours or realisation range.")
    }
    if (!isTRUE(as.logical(z$access_reviewed)) || !isTRUE(as.logical(z$case_mix_reviewed)))
      reason <- c(reason, "Access and comparable case mix must be reviewed.")
    if (nrow(a) != 1L || nrow(b) != 1L) reason <- c(reason, "Both scopes need valid matching capacity and evidence.")
    else {
      if (!a$full_ready || !b$full_ready) reason <- c(reason, "Insufficient evidence or unconfirmed group structure.")
      if (a$clinic_key != b$clinic_key || a$visit_type_key != b$visit_type_key || a$service_format != b$service_format || z$from_key == z$to_key)
        reason <- c(reason, "Compare different time blocks within the same clinic, visit type and service format.")
      if (!length(reason) && (z$hours_to_move > a$staffed_hours - a$protected_min_hours ||
          z$hours_to_move > config$max_hours_move_fraction * a$staffed_hours ||
          z$hours_to_move > z$destination_extra_hours_limit)) reason <- c(reason, "Hours exceed protected capacity, pilot size, or destination limits.")
    }
    if (!length(reason)) {
      lost <- z$hours_to_move * a$completed_per_staffed_hour
      cap <- min(z$additional_patient_demand, z$additional_slots)
      low <- min(cap, z$hours_to_move * b$completed_per_staffed_hour * z$realisation_low) - lost
      high <- min(cap, z$hours_to_move * b$completed_per_staffed_hour * z$realisation_high) - lost
      cost <- z$hours_to_move * (z$destination_incremental_cost_per_hour - z$source_avoidable_cost_per_hour) + z$implementation_cost
      if (cost > z$allowed_budget_increase) reason <- c(reason, "Exceeds the stated incremental budget.")
    }
    status <- if (length(reason)) "Withheld" else if (low > 0) "Candidate for a small pilot" else "No robust gain under supplied assumptions"
    results[[length(results) + 1L]] <- cc_df(scenario = as.character(z$scenario), assessment = status,
      additional_completed_visits_low = low, additional_completed_visits_high = high, incremental_cost = cost,
      explanation = if (length(reason)) paste(reason, collapse = " ") else "Assumption range, not a confidence interval or proven causal effect.",
      next_step = if (length(reason)) "Resolve the listed constraints before considering this option."
        else if (low <= 0) "Retain the current plan unless better evidence or revised feasible assumptions support a gain."
        else "Test one scenario at a time; compare actual completed care, unique patients, access and cost. Do not sum alternative scenarios.")
  }
  list(efficiency = eff, scenarios = cc_bind(results), issues = issues)
}

cc_costed_hours <- function(comparison, plans, config) {
  if (is.null(plans) || !nrow(plans)) return(cc_df(assessment = "No Hours Plans supplied; Basic opening-hours comparisons remain available."))
  req <- c("clinic", "option", "period_start", "period_end", "planned_staffed_hours", "planned_total_cost",
           "budget_limit", "access_reviewed", "duration_reviewed")
  if (length(setdiff(req, names(plans)))) return(cc_df(assessment = paste("Hours Plans missing:", paste(setdiff(req, names(plans)), collapse = ", "))))
  rows <- list()
  keys <- cc_key(plans$clinic, plans$option)
  duplicates <- keys[duplicated(keys) | duplicated(keys, fromLast = TRUE)]
  for (i in seq_len(nrow(plans))) {
    p <- plans[i, ]; e <- if (nrow(comparison)) comparison[comparison$clinic == p$clinic & comparison$option == p$option, , drop = FALSE] else comparison
    reason <- character(); completed <- patients <- cost_per_visit <- NA_real_
    nums <- suppressWarnings(as.numeric(unlist(p[c("planned_staffed_hours", "planned_total_cost", "budget_limit")], use.names = FALSE)))
    dates <- tryCatch(as.Date(c(as.character(p$period_start),as.character(p$period_end))),error=function(e) as.Date(c(NA,NA)))
    if (anyNA(nums) || any(!is.finite(nums)) || any(nums < 0) || nums[1] <= 0) reason <- c(reason,"Missing or invalid staff hours, cost or budget.")
    if (anyNA(dates) || dates[1] != config$as_of-config$lookback_days+1 || dates[2] != config$as_of) reason <- c(reason,"Dates do not match the comparison period.")
    if (keys[i] %in% duplicates || nrow(e) != 1L) reason <- c(reason,"Duplicate or unmatched clinic/option.")
    if (!isTRUE(as.logical(p$access_reviewed)) || !isTRUE(as.logical(p$duration_reviewed))) reason <- c(reason,"Patient access and appointment end times need review.")
    if (nrow(e) == 1L) {
      completed <- e$completed_visits; patients <- e$patients_served
      if (!e$evidence_ready || grepl("unknown", e$interpretation, fixed=TRUE)) reason <- c(reason,"Observed timing evidence is incomplete.")
    }
    if (!anyNA(nums) && all(is.finite(nums)) && all(nums>=0)) {
      cost_per_visit <- cc_ratio(nums[2],completed)
      if (nums[2] > nums[3]) reason <- c(reason,"Above the stated budget.")
    }
    rows[[length(rows)+1L]] <- cc_df(clinic=as.character(p$clinic),option=as.character(p$option),
      captured_completed_visits=completed,captured_patients=patients,planned_staffed_hours=nums[1],
      planned_total_cost=nums[2],budget_limit=nums[3],cost_per_captured_completed_visit=cost_per_visit,
      captured_visits_per_planned_staffed_hour=if(is.finite(nums[1]) && nums[1]>0) completed/nums[1] else NA_real_,
      assessment=if(length(reason)) "Needs review / not eligible" else "Eligible supplied option",
      explanation=if(length(reason)) paste(reason,collapse=" ") else "Historical activity capture under the stated plan; demand and visit times held constant. Not a causal forecast.")
  }
  out <- cc_bind(rows)
  for (clinic in unique(out$clinic)) {
    ok <- which(out$clinic==clinic & out$assessment=="Eligible supplied option")
    if (!length(ok)) next
    # A comparison must use one budget. Conflicting budget assumptions are not ranked.
    if (length(unique(out$budget_limit[ok])) != 1L) {
      out$assessment[ok] <- "Needs review / inconsistent budgets"
      next
    }
    best <- ok[order(-out$captured_completed_visits[ok], -out$captured_patients[ok], out$planned_total_cost[ok])][1]
    out$assessment[best] <- if(length(ok)>1) "Highest captured visits among eligible supplied options" else "Only eligible supplied option; no comparative choice established"
  }
  out
}

cc_analyse <- function(visits, config = cc_defaults(), capacity = NULL, scenarios = NULL, hours_plans = NULL) {
  d <- cc_prepare(visits, config); evidence <- cc_evidence(d, config)
  schedules <- cc_schedule_analysis(d, config)
  recommendations <- cc_recommendations(evidence, schedules, d, config)
  weekly <- d[d$period %in% c("Current", "Previous") & !d$excluded, , drop = FALSE]
  weekly_rows <- if (nrow(weekly)) lapply(split(weekly, with(weekly, cc_key(clinic_key, week))), function(x)
    cc_df(clinic = x$clinic_name[1], week_start = x$week[1], analysed_appointments = nrow(x),
      completed = sum(x$outcome == "Completed"), cancellations = sum(x$outcome == "Cancelled"),
      no_shows = sum(x$outcome == "No-show"))) else list()
  full <- if (config$mode == "full") cc_full(evidence, capacity, scenarios, config) else
    list(efficiency = data.frame(), scenarios = data.frame(), issues = "Basic mode uses only original visit fields; cost and staffed-hour measures are not calculated.")
  full$costed_hours <- if(config$mode=="full") cc_costed_hours(schedules$comparison,hours_plans,config) else data.frame()
  full$input_snapshot <- data.frame()
  if (config$mode == "full") {
    inputs <- list("Hours Plans"=hours_plans, "Capacity"=capacity, "Scenarios"=scenarios)
    snapshot <- list()
    for (table in names(inputs)) {
      input <- inputs[[table]]
      if (is.null(input) || !nrow(input)) next
      for (column in names(input)) snapshot[[length(snapshot)+1L]] <- cc_df(
        input_table=table, input_row=seq_len(nrow(input)), field=column, value=as.character(input[[column]]))
    }
    full$input_snapshot <- cc_bind(snapshot)
  }
  audit <- cc_df(item = c("Mode", "Analysis start", "Analysis end", "Previous period start", "Input rows", "Current rows",
    "Current excluded rows", "Rows with unknown dates", "Rows after analysis end", "Minimum appointments per scope",
    "Minimum observed weeks", "Maximum excluded fraction", "No-show screening threshold", "Cancellation screening threshold",
    "Minimum rate change", "Advance cancellation notice (calendar days)", "Scope of opening-hours model", "Full input status"),
    value = as.character(c(config$mode, format(config$as_of - config$lookback_days + 1), format(config$as_of),
      format(config$as_of - 2 * config$lookback_days + 1), nrow(d), sum(d$period == "Current"),
      sum(d$period == "Current" & d$excluded), sum(d$period == "Unknown date"), sum(d$period == "After analysis end"),
      config$min_appointments, config$min_weeks, config$max_excluded_fraction, config$no_show_threshold,
      config$cancellation_threshold, config$meaningful_rate_change, config$advance_notice_days,
      "Observed start times retained; no new demand, rescheduling, duration, staffing or causal effect inferred.", paste(full$issues, collapse = " "))))
  list(config = config, evidence = evidence, recommendations = recommendations, schedules = schedules,
       weekly = cc_bind(weekly_rows), full = full, audit = audit)
}

cc_tables <- function(result) {
  tables <- list("Current Activity" = result$schedules$scopes,
    "Schedule Comparison" = result$schedules$comparison, "Recommendations" = result$recommendations,
    "Decision Evidence" = result$evidence, "Weekly Evidence" = result$weekly,
    "Inputs and Rules" = result$audit)
  if (result$config$mode == "full") tables <- c(tables, list("Costed Hours" = result$full$costed_hours, "Efficiency" = result$full$efficiency,
    "Full Scenarios" = result$full$scenarios, "Planning Inputs" = result$full$input_snapshot))
  tables
}
