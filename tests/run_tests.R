source("R/decision_support.R")
passed <- 0L
check <- function(label, condition) {
  if (!isTRUE(condition)) stop("FAILED: ", label)
  passed <<- passed + 1L; cat("PASS", label, "\n")
}
expect_error <- function(expr) inherits(tryCatch({force(expr);NULL},error=identity),"error")
make_visits <- function(seconds, status=rep("comp",length(seconds)), patients=paste0("P",seq_along(seconds))) {
  n <- length(seconds)
  cc_df(visit_record_id=seq_len(n), mrn=patients, clinic_key="C1", clinic_name="Synthetic clinic",
    visit_date=rep(as.Date("2026-08-24"),n), appointment_time_seconds=seconds,
    visit_type_grouping_key="T1", internal_visit_type_name="Follow-up", appt_status_standardized=status,
    cancellation_date=rep(as.Date(NA),n), possible_duplicate_appointment=FALSE,
    data_quality_issue=NA_character_, potential_group_visit=FALSE, session_id=paste0("S",seq_len(n)))
}
cfg <- cc_defaults();cfg$as_of <- as.Date("2026-08-30");cfg$min_appointments <- 1L;cfg$min_weeks <- 1L
d <- make_visits(c(7*3600+59*60,8*3600,9*3600,15*3600+59*60,16*3600,16.5*3600,17*3600),
                 patients=c("P0","P1","P2","P3","P1","P4","P5"))
x <- cc_analyse(d,cfg); s <- x$schedules$comparison
check("opening inclusive, closing exclusive", identical(s$completed_visits,c(3L,5L,4L)))
check("distinct patients recomputed per window", identical(s$patients_served,c(3L,4L,4L)))
check("all alternatives use the same denominator", all(s$comparison_days==1L))
check("8-5 is nine hours, equal spans are eight", identical(s$window_hours,c(8,9,8)))
check("reference deltas reconcile", identical(s$change_vs_reference,c(0L,2L,1L)))
check("added and removed completed visits reconcile", all(s$completed_added-s$completed_removed==s$change_vs_reference))
check("lost-patient counts are distinct", s$patients_losing_all_observed_completed_visits[2]==2L)
check("Basic never computes cost even when extra inputs supplied", nrow(x$full$efficiency)==0L)

d2 <- d;d2$possible_duplicate_appointment[2] <- TRUE; d2$appt_status_standardized[3] <- "Scheduled"
r <- cc_analyse(d2,cfg)
check("duplicates and unresolved status excluded from decisions", r$schedules$comparison$completed_visits[1]==1L)
check("exclusion audit reconciles", r$audit$value[r$audit$item=="Current excluded rows"]=="2")
check("poor evidence does not recommend moving hours", !any(r$recommendations$rule %in% c("HR01","HR02")))
d3 <- d;d3$visit_record_id[2] <- d3$visit_record_id[1]
check("provider-attributed rows rejected", expect_error(cc_analyse(d3,cfg)))
d4 <- d; d4$visit_date[1] <- cfg$as_of+1;d4$visit_date[2] <- NA
r <- cc_analyse(d4,cfg)
check("future and undated records audited", r$audit$value[r$audit$item=="Rows after analysis end"]=="1" && r$audit$value[r$audit$item=="Rows with unknown dates"]=="1")
r <- cc_analyse(make_visits(c(9,10,11)*3600),cfg)
check("absence of activity in extension is unknown demand", grepl("unknown",r$schedules$comparison$interpretation[2]))
bad <- cfg;bad$schedules$end[1] <- bad$schedules$start[1]
check("invalid schedules rejected",expect_error(cc_analyse(d,bad)))
bad <- cfg;bad$min_appointments <- -1
check("invalid screening thresholds rejected",expect_error(cc_analyse(d,bad)))
d5 <- d;d5$visit_date <- as.Date("2025-01-01")
check("empty current period is handled",nrow(cc_analyse(d5,cfg)$schedules$comparison)==0L)
d6 <- make_visits(c(8,9,16)*3600,c("can","no show","comp"));d6$cancellation_date[1] <- as.Date("2026-08-20")
r <- cc_analyse(d6,cfg)
check("cancellations and no-shows stay separate",r$schedules$comparison$cancellations[2]==1L && r$schedules$comparison$no_shows[2]==1L)
check("advance cancellations use date difference",sum(r$evidence$advance_cancellations)==1L)

# Independent Full scenario fixture: source rate 1/h, destination 3/h.
ev <- cc_df(period=c("Current","Current"),clinic_key="C1",clinic="Synthetic clinic",visit_type_key="T1",
  visit_type="Follow-up",service_format="Individual",time_block=c("8-9","9-16"),source_records=c(80,120),
  excluded_records=0L,analysed_appointments=c(80,120),completed=c(40,120),patients_served=c(40,120),
  cancellations=0L,no_shows=c(40,0),advance_cancellations=0L,weeks_observed=8L,
  completion_rate=c(0.5,1),cancellation_rate=0,no_show_rate=c(0.5,0),excluded_fraction=0,evidence_ready=TRUE,evidence_key=c("FROM","TO"))
cap <- cc_df(evidence_key=c("FROM","TO"),period_start="2026-07-06",period_end="2026-08-30",
  staffed_hours=40,bookable_patient_slots=160,allocated_cost=2000,protected_min_hours=20,group_confirmed=FALSE)
sc <- cc_df(scenario="Pilot",from_key="FROM",to_key="TO",hours_to_move=4,additional_patient_demand=20,
  additional_slots=20,destination_extra_hours_limit=8,realisation_low=0.5,realisation_high=1,
  source_avoidable_cost_per_hour=0,destination_incremental_cost_per_hour=0,implementation_cost=0,
  allowed_budget_increase=0,access_reviewed=TRUE,case_mix_reviewed=TRUE)
r <- cc_full(ev,cap,sc,cfg)
check("Full gain subtracts service lost at source",r$scenarios$additional_completed_visits_low==2 && r$scenarios$additional_completed_visits_high==8)
check("allocated cost never becomes automatic cash savings",r$scenarios$incremental_cost==0)
sc2 <- sc;sc2$additional_patient_demand <- 1
check("unmet demand caps modelled additional care",cc_full(ev,cap,sc2,cfg)$scenarios$additional_completed_visits_high == -3)
sc2 <- sc;sc2$additional_slots <- 2
check("available slots cap modelled additional care",cc_full(ev,cap,sc2,cfg)$scenarios$additional_completed_visits_high == -2)
sc2 <- sc;sc2$destination_incremental_cost_per_hour <- 1
check("budget breach withholds recommendation",cc_full(ev,cap,sc2,cfg)$scenarios$assessment=="Withheld")
sc2 <- sc;sc2$access_reviewed <- FALSE
check("unreviewed access blocks scenario",cc_full(ev,cap,sc2,cfg)$scenarios$assessment=="Withheld")
sc2 <- sc;sc2$realisation_low <- 1.1
check("impossible realisation range rejected",cc_full(ev,cap,sc2,cfg)$scenarios$assessment=="Withheld")
cap2 <- cap;cap2$period_end <- "2026-08-29"
check("wrong-period denominators withheld",nrow(cc_full(ev,cap2,sc,cfg)$efficiency)==0)
cap2 <- rbind(cap,cap[1,])
check("duplicate capacity scope not double counted",!"FROM" %in% cc_full(ev,cap2,sc,cfg)$efficiency$evidence_key)
cap2 <- cap;cap2$staffed_hours[1] <- 0
check("zero staff hours do not produce infinity",cc_full(ev,cap2,sc,cfg)$scenarios$assessment=="Withheld")
cap2 <- cap;cap2$protected_min_hours[1] <- 39
check("protected service floor is enforced",cc_full(ev,cap2,sc,cfg)$scenarios$assessment=="Withheld")
ev2 <- ev;ev2$service_format <- "Inferred group"
check("unconfirmed inferred groups cannot drive full decisions",cc_full(ev2,cap,sc,cfg)$scenarios$assessment=="Withheld")
check("Full inputs missing preserve Basic path",length(cc_full(ev,NULL,NULL,cfg)$issues)==1L)
plans <- cc_df(clinic=s$clinic,option=s$option,period_start="2026-07-06",period_end="2026-08-30",
  planned_staffed_hours=c(8,9,8),planned_total_cost=c(800,900,800),budget_limit=800,
  access_reviewed=TRUE,duration_reviewed=TRUE)
r <- cc_costed_hours(s,plans,cfg)
check("costed 8-5 does not win by using extra budget",r$assessment[2]=="Needs review / not eligible")
check("costed equal-span choice uses captured activity",grepl("Highest",r$assessment[3]))
check("costed hours use planned staffing denominator",r$captured_visits_per_planned_staffed_hour[3]==0.5)
plans2 <- plans;plans2$duration_reviewed <- FALSE
check("unknown end-of-visit feasibility blocks costed choice",all(cc_costed_hours(s,plans2,cfg)$assessment=="Needs review / not eligible"))
plans2 <- plans;plans2$budget_limit[3] <- 801
check("different budgets cannot establish a preferred hours option",!any(grepl("Highest",cc_costed_hours(s,plans2,cfg)$assessment)))
plans2 <- plans;plans2$planned_staffed_hours <- NA_real_
check("missing planned staff hours remain missing",all(cc_costed_hours(s,plans2,cfg)$assessment=="Needs review / not eligible"))
plans2 <- plans;plans2$period_end <- "2026-08-29"
check("costed options require matching dates",all(cc_costed_hours(s,plans2,cfg)$assessment=="Needs review / not eligible"))
cat("\n",passed," decision checks passed.\n",sep="")
