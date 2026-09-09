# Shiny runs this directory; the shared engine stays one directory above it.
project_root <- normalizePath("..", mustWork = TRUE)
app_env <- new.env(parent = globalenv())
previous <- options(clinic_capacity_no_autorun = TRUE)
sys.source(file.path(project_root, "run_report.R"), envir = app_env)
options(previous)
for (file in c("R/decision_support.R", "R/workbook_support.R", "app/helpers.R", "app/ui.R", "app/server.R"))
  sys.source(file.path(project_root, file), envir = app_env)
shiny::shinyApp(app_env$cc_app_ui(), app_env$cc_app_server(project_root))
