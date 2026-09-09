# Run with Rscript run_app.R or source("run_app.R") from the project folder.
cc_launch_app <- function(project_root = NULL, launch_browser = interactive()) {
  if (is.null(project_root)) {
    script <- grep("^--file=", commandArgs(), value = TRUE)
    project_root <- if (length(script)) dirname(normalizePath(sub("^--file=", "", script[1]))) else getwd()
  }
  if (!requireNamespace("shiny", quietly = TRUE))
    stop("The Interface Version needs the shiny R package. See docs/INTERFACE_SETUP.md. The Hospital R Tool works without it.", call. = FALSE)
  options(shiny.maxRequestSize = 200 * 1024^2, shiny.sanitize.errors = TRUE)
  shiny::runApp(file.path(normalizePath(project_root), "app"), host = "127.0.0.1",
    launch.browser = launch_browser, display.mode = "normal")
}
if (!isTRUE(getOption("clinic_capacity_no_autorun", FALSE))) cc_launch_app(launch_browser = TRUE)
