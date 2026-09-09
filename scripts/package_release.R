# Build both deliverables from one tracked source tree. Never package data folders.
args <- commandArgs(trailingOnly = TRUE)
root <- normalizePath(getwd())
if (!file.exists("run_report.R") || !dir.exists("app")) stop("Run from the development project folder.")
if (!requireNamespace("zip", quietly = TRUE)) stop("Install the zip package used by openxlsx.")
destination <- if (length(args)) args[1] else "releases"
dir.create(destination, recursive = TRUE, showWarnings = FALSE)
destination <- normalizePath(destination)
tracked <- system2("git", c("ls-files"), stdout = TRUE)
if (!length(tracked) || !is.null(attr(tracked, "status"))) stop("A Git checkout with tracked source files is required.")
if (any(grepl("^(input_data|output)/", tracked) & !grepl("/\\.gitkeep$", tracked)))
  stop("Input or output data is tracked. Remove it from the release before packaging.")
blocked <- grepl("(^|/)(\\.env($|\\.)|\\.Rhistory$|\\.RData$)|\\.[rR][dD][sS]$", tracked)
if (any(blocked)) stop("Private runtime files are tracked; inspect before packaging.")
hospital <- c("LICENSE", "config.R", "run_report.R", "Run-Basic.cmd", "Run-Full.cmd", "Clinic-Capacity-Planner.Rproj",
  "R/decision_support.R", "R/workbook_support.R", "R/report_pipeline.R", "START-HOSPITAL.md",
  "docs/HOSPITAL_SETUP.md", "docs/FULL_INPUTS.md", "docs/ACCEPTANCE.md", "input_data/.gitkeep", "output/.gitkeep")
if (length(setdiff(hospital, tracked))) stop("Hospital release files must be tracked in Git.")
stage <- tempfile("cc_releases_"); dir.create(stage)
package <- function(files, directory_name, archive_name) {
  folder <- file.path(stage, directory_name); dir.create(folder)
  for (file in files) {
    dest <- file.path(folder, file)
    dir.create(dirname(dest), recursive = TRUE, showWarnings = FALSE)
    if (!file.copy(file.path(root, file), dest)) stop("Cannot stage ", file)
    if (grepl("[.]cmd$", file, ignore.case = TRUE))
      writeBin(charToRaw(paste0(paste(readLines(dest, warn = FALSE), collapse = "\r\n"), "\r\n")), dest)
  }
  zip::zipr(file.path(destination, archive_name), directory_name, root = stage)
  cat(file.path(destination, archive_name), "\n")
}
tryCatch({
  package(hospital, "Clinic-Capacity-Hospital-R", "Clinic-Capacity-Hospital-R.zip")
  package(tracked, "clinic-capacity-planner", "Clinic-Capacity-Planner-source.zip")
}, finally = unlink(stage, recursive = TRUE))
