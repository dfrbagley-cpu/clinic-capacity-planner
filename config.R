# Edit this file in a text editor or RStudio. All paths may be shared-drive paths.
settings <- list(
  mode = "basic",                         # "basic" or "full"
  input_directory = "input_data",
  output_directory = "output",
  protected = TRUE,                       # original password-protected Excel exports
  input_sheet = 1,
  date_order = "mdy",                      # explicit locale for ambiguous text dates
  as_of = Sys.Date() - 1,                  # last day included in current analysis
  lookback_days = 56L,                     # prior comparison uses another 56 days
  full_inputs = "planning_inputs.xlsx"    # optional; never required by Basic
)
