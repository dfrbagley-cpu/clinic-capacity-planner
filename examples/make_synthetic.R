# Fictional appointments with the EXACT 12 source columns required by V3.
cc_make_synthetic <- function(path) {
  rows <- list(); n <- 0L
  add <- function(clinic, provider, type, day, time, patient, status, notice = NA_real_) {
    n <<- n + 1L
    rows[[n]] <<- data.frame(MRN = patient, Patient = paste("Synthetic patient", patient),
      Department = clinic, Dept = "Synthetic staff department [SD1]", `Provider/Resource` = provider,
      `Visit Type` = type, Type = type, `Visit Date` = as.Date(day), Time = time / 24,
      `Appt Status` = status, `Canc Date` = if (is.na(notice)) as.Date(NA) else as.Date(day) - notice,
      `Canc Reason` = if (is.na(notice)) NA_character_ else "Synthetic patient cancellation",
      check.names = FALSE, stringsAsFactors = FALSE)
  }
  for (w in 0:15) {
    monday <- as.Date("2026-05-11") + 7 * w
    for (clinic_num in 1:2) {
      clinic <- c("Meridian Clinic [MC1]", "Cedar Clinic [CC2]")[clinic_num]
      day <- monday + clinic_num - 1
      for (j in 0:3) {
        status <- if (clinic_num == 1 && j < 2) "No Show" else "Comp"
        add(clinic, "Synthetic clinician A [SP1]", "Follow-up [FU1]", day, 8 + j/4,
            paste0("SYN-",clinic_num,"-EARLY-",w,"-",j), status)
      }
      for (j in 0:9) {
        provider <- if (j == 4) "Synthetic clinician A [SP1]; Synthetic clinician B [SP2]" else "Synthetic clinician A [SP1]"
        add(clinic, provider, "Follow-up [FU1]", day, 9 + j/2,
            paste0("SYN-",clinic_num,"-CORE-",w,"-",j), "Comp")
      }
      for (j in 0:5) {
        status <- if (clinic_num == 1) "Comp" else if (j < 4) "Can" else "No Show"
        add(clinic, "Synthetic clinician B [SP2]", "Follow-up [FU1]", day, 16 + j/6,
            paste0("SYN-",clinic_num,"-LATE-",w,"-",j), status, if (status == "Can") 3 else NA_real_)
      }
      for (j in 1:3) add(clinic, "Synthetic group lead [SG1]; Synthetic co-lead [SG2]", "Group care [GC1]", day, 14,
        paste0("SYN-",clinic_num,"-GROUP-",w,"-",j), if (j < 3) "Comp" else "No Show")
    }
  }
  raw <- do.call(rbind, rows)
  # Exercise quality control without changing the original input schema.
  raw <- rbind(raw, raw[nrow(raw)-8, , drop=FALSE])
  add("Meridian Clinic [MC1]", "Synthetic clinician A [SP1]", "Follow-up [FU1]", as.Date("2026-09-02"), 10, "SYN-FUTURE", "Scheduled")
  raw <- rbind(raw, rows[[n]])
  dir.create(dirname(path), recursive=TRUE, showWarnings=FALSE)
  openxlsx::write.xlsx(raw, path, overwrite=TRUE, sheetName="Visits")
  invisible(raw)
}
