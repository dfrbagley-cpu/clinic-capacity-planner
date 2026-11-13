# Capture a local runtime inventory, without installing packages or recording paths.
# This is provenance, not a lockfile or a claim of validated package compatibility.
args <- commandArgs(trailingOnly = TRUE)
if (length(args) != 1L || !nzchar(args[1L]))
  stop("Usage: Rscript scripts/capture_environment.R <new-environment.csv>")
if (file.exists(args[1L])) stop("Refusing to overwrite an existing inventory.")
packages <- installed.packages()
records <- data.frame(component = c("R", "R-platform", packages[, "Package"]),
  version = c(as.character(getRversion()), R.version$platform, packages[, "Version"]),
  kind = c("runtime", "platform", rep("package", nrow(packages))),
  stringsAsFactors = FALSE)
records <- unique(records)
records <- records[order(records$kind, records$component, records$version), ]
write.csv(records, args[1L], row.names = FALSE, fileEncoding = "UTF-8")
cat("Local environment inventory saved. Review before sharing.\n")
