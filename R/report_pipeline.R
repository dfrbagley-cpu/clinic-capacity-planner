# ==============================================================================
# CLINIC CAPACITY PLANNER - HOSPITAL REPORT PIPELINE
# ==============================================================================
#
# Purpose:
#   Import monthly protected CPC visit workbooks, combine the data, analyze
#   clinic scheduling activity, attribute joint visits to individual providers,
#   and create a manager-ready Excel report with Basic/Full decision support.
#
# Project requirements:
#   - Open the CPC Visit Analysis R project before running this script.
#   - Place all monthly source workbooks in the input_data folder.
#   - All source workbooks must use the same password.
#   - Microsoft Excel desktop must be installed.
#   - Close the source workbooks before running the tool.
#
# Privacy:
#   MRN and Patient are used internally for analysis but are excluded from the
#   output workbook.
#
# ==============================================================================


# ==============================================================================
# 1. LOAD AND VALIDATE REQUIRED PACKAGES
# ==============================================================================
#
# Confirms that every required R package is installed and then loads the
# packages used for importing, cleaning, analyzing, and exporting the data.
#
# ==============================================================================

required_packages <- c(
    "dplyr",
  "purrr",
  "stringr",
  "lubridate",
  "openxlsx",
  "tidyr",
  "tibble"
)

if (settings$protected) required_packages <- c(required_packages, "excel.link")

installed_package_names <- rownames(
  installed.packages()
)

missing_packages <- setdiff(
  required_packages,
  installed_package_names
)

if (length(missing_packages) > 0) {

  stop(
    paste0(
      "The following required packages are not installed:\n\n",
      paste(
        missing_packages,
        collapse = ", "
      ),
      "\n\nRun this command once:\n\n",
      "install.packages(c(",
      paste0(
        '"',
        missing_packages,
        '"',
        collapse = ", "
      ),
      "))"
    ),
    call. = FALSE
  )
}

if (settings$protected) library(excel.link)
library(dplyr)
library(purrr)
library(stringr)
library(lubridate)
library(openxlsx)
library(tidyr)
library(tibble)


# ==============================================================================
# 2. DEFINE PROJECT FOLDERS
# ==============================================================================
#
# Uses paths relative to the R project. The script expects input files in
# input_data and saves completed reports in output.
#
# ==============================================================================

input_directory <- settings$input_directory
output_directory <- settings$output_directory

if (!dir.exists(input_directory)) {

  stop(
    paste0(
      "The input_data folder was not found.\n\n",
      "Open the CPC Visit Analysis R project before running this script.\n\n",
      "Current working directory:\n",
      normalizePath(
        ".",
        winslash = "/",
        mustWork = FALSE
      )
    ),
    call. = FALSE
  )
}

dir.create(
  path = output_directory,
  showWarnings = FALSE,
  recursive = TRUE
)


# ==============================================================================
# 3. DEFINE REQUIRED SOURCE COLUMNS
# ==============================================================================
#
# Lists the columns expected in every monthly workbook.
#
# Department identifies the appointment clinic/program and is imported as
# Appt Department. Dept identifies the staff member's organizational department
# and is imported as Provider Department. Provider Department is retained but is
# not currently used in any calculation or grouping.
#
# ==============================================================================

required_columns <- c(
  "MRN",
  "Patient",
  "Department",
  "Dept",
  "Provider/Resource",
  "Visit Type",
  "Type",
  "Visit Date",
  "Time",
  "Appt Status",
  "Canc Date",
  "Canc Reason"
)


# ==============================================================================
# 4. REQUEST THE SHARED WORKBOOK PASSWORD
# ==============================================================================
#
# Prompts once for the password shared by all source workbooks. The password is
# held temporarily in the current R session and is not written to the report.
#
# ==============================================================================

if (settings$protected) {
  workbook_password <- cc_password()
  if (is.null(workbook_password) || length(workbook_password) != 1L || is.na(workbook_password) || !nzchar(workbook_password))
    stop("No workbook password was entered.", call. = FALSE)
}


# ==============================================================================
# 5. DEFINE DATA-CLEANING AND PARSING FUNCTIONS
# ==============================================================================
#
# Creates reusable functions for:
#   - Converting blank text to missing values
#   - Reading Excel dates and times
#   - Extracting bracketed identifiers
#   - Parsing individual providers from joint provider fields
#   - Counting unique visits safely
#
# ==============================================================================


# ------------------------------------------------------------------------------
# 5.1 Convert blank text to missing values
# ------------------------------------------------------------------------------

blank_to_na <- function(x) {

  x <- as.character(x)

  x <- stringr::str_squish(
    x
  )

  x[
    is.na(x) |
      x %in% c(
        "",
        "NA",
        "N/A",
        "NULL"
      )
  ] <- NA_character_

  x
}


# ------------------------------------------------------------------------------
# 5.2 Convert Excel or text values to dates
# ------------------------------------------------------------------------------

parse_excel_date <- function(x) {

  if (inherits(x, "Date")) {
    return(as.Date(x))
  }

  if (inherits(x, "POSIXt")) {
    return(as.Date(x))
  }

  if (is.numeric(x)) {

    return(
      as.Date(
        x,
        origin = "1899-12-30"
      )
    )
  }

  x_character <- blank_to_na(
    x
  )

  numeric_values <- suppressWarnings(
    as.numeric(x_character)
  )

  result <- rep(
    as.Date(NA),
    length(x_character)
  )

  excel_serial_rows <- !is.na(numeric_values) &
    numeric_values > 1000

  if (any(excel_serial_rows)) {

    result[excel_serial_rows] <- as.Date(
      numeric_values[excel_serial_rows],
      origin = "1899-12-30"
    )
  }

  text_date_rows <- !excel_serial_rows &
    !is.na(x_character)

  if (any(text_date_rows)) {

    parsed_dates <- suppressWarnings(
      lubridate::parse_date_time(
        x_character[text_date_rows],
        orders = c(
          "ymd",
          settings$date_order,
          "Ymd HMS",
          "Ymd HM",
          paste(settings$date_order, "HMS"),
          paste(settings$date_order, "HM")
        ),
        quiet = TRUE
      )
    )

    result[text_date_rows] <- as.Date(
      parsed_dates
    )
  }

  result
}


# ------------------------------------------------------------------------------
# 5.3 Convert Excel or text appointment times to seconds after midnight
# ------------------------------------------------------------------------------

parse_time_seconds <- function(x) {

  if (inherits(x, "POSIXt")) {

    return(
      lubridate::hour(x) * 3600 +
        lubridate::minute(x) * 60 +
        floor(
          lubridate::second(x)
        )
    )
  }

  if (is.numeric(x)) {

    return(
      round(
        (x %% 1) * 86400
      )
    )
  }

  x_character <- blank_to_na(
    x
  )

  result <- rep(
    NA_real_,
    length(x_character)
  )

  numeric_values <- suppressWarnings(
    as.numeric(x_character)
  )

  numeric_time_rows <- !is.na(numeric_values) &
    numeric_values >= 0 &
    numeric_values < 1

  if (any(numeric_time_rows)) {

    result[numeric_time_rows] <- round(
      numeric_values[numeric_time_rows] * 86400
    )
  }

  text_time_rows <- !numeric_time_rows &
    !is.na(x_character)

  if (any(text_time_rows)) {

    parsed_times <- suppressWarnings(
      lubridate::parse_date_time(
        x_character[text_time_rows],
        orders = c(
          "I:M:S p",
          "I:M p",
          "H:M:S",
          "H:M"
        ),
        quiet = TRUE
      )
    )

    result[text_time_rows] <- ifelse(
      is.na(parsed_times),
      NA_real_,
      lubridate::hour(parsed_times) * 3600 +
        lubridate::minute(parsed_times) * 60 +
        floor(
          lubridate::second(parsed_times)
        )
    )
  }

  result
}


# ------------------------------------------------------------------------------
# 5.4 Convert seconds after midnight to a readable 12-hour time
# ------------------------------------------------------------------------------

format_time_12_hour <- function(
    seconds_after_midnight
) {

  result <- rep(
    NA_character_,
    length(seconds_after_midnight)
  )

  valid_rows <- !is.na(
    seconds_after_midnight
  )

  if (!any(valid_rows)) {
    return(result)
  }

  total_seconds <- round(
    seconds_after_midnight[valid_rows]
  )

  hour_24 <- floor(
    total_seconds / 3600
  ) %% 24

  minute_value <- floor(
    (total_seconds %% 3600) / 60
  )

  hour_12 <- hour_24 %% 12

  hour_12[
    hour_12 == 0
  ] <- 12

  am_pm <- ifelse(
    hour_24 < 12,
    "AM",
    "PM"
  )

  result[valid_rows] <- paste0(
    hour_12,
    ":",
    sprintf(
      "%02d",
      minute_value
    ),
    " ",
    am_pm
  )

  result
}


# ------------------------------------------------------------------------------
# 5.5 Extract an identifier from square brackets
# ------------------------------------------------------------------------------

extract_bracket_id <- function(x) {

  x <- blank_to_na(
    x
  )

  extracted_id <- stringr::str_match(
    x,
    "\\[([^\\[\\]]+)\\]\\s*$"
  )[, 2]

  blank_to_na(
    extracted_id
  )
}


# ------------------------------------------------------------------------------
# 5.6 Remove a bracketed identifier from a name
# ------------------------------------------------------------------------------

remove_bracket_id <- function(x) {

  x <- blank_to_na(
    x
  )

  cleaned_name <- stringr::str_remove(
    x,
    "\\s*\\[[^\\[\\]]+\\]\\s*$"
  )

  blank_to_na(
    cleaned_name
  )
}


# ------------------------------------------------------------------------------
# 5.7 Create a stable grouping key
# ------------------------------------------------------------------------------

create_grouping_key <- function(
    id,
    name,
    prefix
) {

  dplyr::case_when(
    !is.na(id) ~ paste0(
      prefix,
      "_ID_",
      id
    ),

    !is.na(name) ~ paste0(
      prefix,
      "_NAME_",
      stringr::str_to_upper(
        stringr::str_squish(name)
      )
    ),

    TRUE ~ NA_character_
  )
}


# ------------------------------------------------------------------------------
# 5.8 Count distinct visit records meeting a condition
# ------------------------------------------------------------------------------

safe_distinct_count <- function(
    record_ids,
    condition
) {

  records_to_count <- record_ids[
    !is.na(condition) &
      condition
  ]

  dplyr::n_distinct(
    records_to_count,
    na.rm = TRUE
  )
}


# ------------------------------------------------------------------------------
# 5.9 Extract individual providers from a joint Provider/Resource field
# ------------------------------------------------------------------------------

extract_provider_entries <- function(
    provider_value
) {

  provider_value <- blank_to_na(
    provider_value
  )

  if (
    length(provider_value) == 0 ||
    is.na(provider_value)
  ) {
    return(NA_character_)
  }

  provider_entries <- stringr::str_extract_all(
    provider_value,
    "[^\\[]+\\[[^\\]]+\\]"
  )[[1]]

  if (length(provider_entries) == 0) {

    return(
      stringr::str_squish(
        provider_value
      )
    )
  }

  provider_entries <- stringr::str_remove(
    provider_entries,
    "^\\s*[,;|/]+\\s*"
  )

  provider_entries <- stringr::str_squish(
    provider_entries
  )

  provider_entries <- provider_entries[
    !is.na(provider_entries) &
      provider_entries != ""
  ]

  unique(
    provider_entries
  )
}


# ==============================================================================
# 6. IDENTIFY EXCEL INPUT FILES
# ==============================================================================
#
# Finds every Excel workbook in input_data. Temporary Excel lock files beginning
# with ~$ are excluded automatically.
#
# ==============================================================================

input_files <- list.files(
  path = input_directory,
  pattern = "\\.(xlsx|xls)$",
  full.names = TRUE,
  ignore.case = TRUE
)

input_files <- input_files[
  !startsWith(
    basename(input_files),
    "~$"
  )
]

input_files <- sort(
  input_files
)

if (length(input_files) == 0) {

  stop(
    paste0(
      "No Excel files were found in:\n",
      normalizePath(
        input_directory,
        winslash = "/",
        mustWork = FALSE
      )
    ),
    call. = FALSE
  )
}

message("")
message(
  "Excel input files found: ",
  length(input_files)
)

purrr::walk(
  input_files,
  function(file_path) {

    message(
      "  - ",
      basename(file_path)
    )
  }
)


# ==============================================================================
# 7. IMPORT AND VALIDATE ONE PASSWORD-PROTECTED WORKBOOK
# ==============================================================================
#
# Opens one workbook through Microsoft Excel, validates its structure, corrects
# generic Excel column names when necessary, and standardizes the source names.
#
# ==============================================================================

read_cpc_visit_file <- function(
    file_path
) {

  source_file_name <- basename(
    file_path
  )

  message(
    "Importing workbook: ",
    source_file_name
  )

  if (!file.exists(file_path)) {

    stop(
      paste0(
        "The input file could not be found:\n",
        file_path
      ),
      call. = FALSE
    )
  }

  file_details <- file.info(
    file_path
  )

  if (
    is.na(file_details$size) ||
    file_details$size == 0
  ) {

    stop(
      paste0(
        "The input file is empty or cannot be read:\n",
        file_path
      ),
      call. = FALSE
    )
  }

  absolute_file_path <- normalizePath(
    file_path,
    winslash = "\\",
    mustWork = TRUE
  )

  imported_data <- tryCatch(
    {
      if (!settings$protected) {
        openxlsx::read.xlsx(file_path, sheet = settings$input_sheet, detectDates = TRUE,
                            skipEmptyRows = FALSE, check.names = FALSE, sep.names = " ")
      } else excel.link::xl.read.file(
        filename = absolute_file_path,
        header = TRUE,
        row.names = FALSE,
        col.names = NULL,
        xl.sheet = settings$input_sheet,
        top.left.cell = "A1",
        na = "",
        password = workbook_password,
        write.res.password = NULL,
        excel.visible = FALSE
      )
    },
    error = function(e) {

      stop(
        paste0(
          "The protected workbook could not be imported:\n\n",
          source_file_name,
          "\n\nPossible reasons:\n",
          "- The workbook password is incorrect.\n",
          "- The workbook is open or locked.\n",
          "- Microsoft Excel desktop is unavailable.\n",
          "- The visit worksheet was not active when saved.\n",
          "- The headers do not begin in cell A1.\n\n",
          "Technical message:\n",
          conditionMessage(e)
        ),
        call. = FALSE
      )
    }
  )

  imported_data <- as.data.frame(
    imported_data,
    stringsAsFactors = FALSE,
    check.names = FALSE
  )

  if (nrow(imported_data) == 0) {

    stop(
      paste0(
        "No visit rows were found in:\n",
        source_file_name
      ),
      call. = FALSE
    )
  }

  nonempty_rows <- apply(
    imported_data,
    MARGIN = 1,
    FUN = function(row_values) {

      row_values <- as.character(
        row_values
      )

      any(
        !is.na(row_values) &
          stringr::str_squish(row_values) != ""
      )
    }
  )

  imported_data <- imported_data[
    nonempty_rows,
    ,
    drop = FALSE
  ]

  names(imported_data) <- stringr::str_squish(
    names(imported_data)
  )

  if (nrow(imported_data) > 0) {

    first_row_values <- imported_data[1, ] |>
      unlist(
        use.names = FALSE
      ) |>
      as.character() |>
      stringr::str_squish()

    if (
      length(first_row_values) == length(required_columns) &&
      setequal(
        first_row_values,
        required_columns
      )
    ) {

      message(
        "  Header row found in the first imported row."
      )

      names(imported_data) <- first_row_values

      imported_data <- imported_data[
        -1,
        ,
        drop = FALSE
      ]
    }
  }

  expected_generic_names <- letters[
    seq_len(
      min(
        ncol(imported_data),
        length(letters)
      )
    )
  ]

  generic_columns_detected <- (
    ncol(imported_data) ==
      length(required_columns) &&
      identical(
        tolower(
          names(imported_data)
        ),
        expected_generic_names[
          seq_len(
            ncol(imported_data)
          )
        ]
      )
  )

  if (generic_columns_detected) {
    stop("Source headers must be present; columns are not assigned by position.", call. = FALSE)
  }

  missing_columns <- setdiff(
    required_columns,
    names(imported_data)
  )

  if (length(missing_columns) > 0) {

    stop(
      paste0(
        "Required columns are missing from:\n",
        source_file_name,
        "\n\nMissing columns:\n",
        paste(
          missing_columns,
          collapse = ", "
        ),
        "\n\nColumns found:\n",
        paste(
          names(imported_data),
          collapse = ", "
        )
      ),
      call. = FALSE
    )
  }

  imported_data <- imported_data |>
    dplyr::select(
      dplyr::all_of(
        required_columns
      )
    )

  imported_data |>
    dplyr::mutate(
      source_file =
        source_file_name,

      source_row =
        dplyr::row_number() + 1
    ) |>
    dplyr::rename(
      mrn =
        MRN,

      patient =
        Patient,

      appt_department =
        Department,

      provider_department =
        Dept,

      provider_resource =
        `Provider/Resource`,

      visit_type =
        `Visit Type`,

      internal_type =
        Type,

      visit_date_raw =
        `Visit Date`,

      appointment_time_raw =
        Time,

      appt_status =
        `Appt Status`,

      cancellation_date_raw =
        `Canc Date`,

      cancellation_reason =
        `Canc Reason`
    )
}


# ==============================================================================
# 8. IMPORT AND COMBINE ALL MONTHLY WORKBOOKS
# ==============================================================================
#
# Runs the import function for every workbook and combines the monthly data into
# one dataset. A unique Visit Record ID is assigned to each imported row.
#
# ==============================================================================

combined_visits <- purrr::map_dfr(
  input_files,
  read_cpc_visit_file
)

if (nrow(combined_visits) == 0) {

  stop(
    "The workbooks were imported, but no visit rows were found.",
    call. = FALSE
  )
}

combined_visits <- combined_visits |>
  dplyr::mutate(
    visit_record_id =
      dplyr::row_number()
  )

message(
  "Total visit rows imported: ",
  format(
    nrow(combined_visits),
    big.mark = ","
  )
)


# ==============================================================================
# END OF VERSION 3 - BASE IMPORT SECTION
# Paste Section 2 immediately below this line.
# ==============================================================================

# ==============================================================================
# VERSION 3 - BASE ANALYSIS SECTION
# ==============================================================================


# ==============================================================================
# 9. CLEAN AND STANDARDIZE THE COMBINED VISIT DATA
# ==============================================================================
#
# Standardizes text, dates, appointment times, reporting months, appointment
# departments, provider departments, visit types, and grouping identifiers.
#
# Clinic reporting is derived only from Appt Department. Provider Department is
# retained for reference but is not used in calculations.
#
# ==============================================================================

visits <- combined_visits |>
  dplyr::mutate(
    mrn =
      blank_to_na(mrn),

    patient =
      blank_to_na(patient),

    appt_department =
      blank_to_na(appt_department),

    provider_department =
      blank_to_na(provider_department),

    provider_resource =
      blank_to_na(provider_resource),

    visit_type =
      blank_to_na(visit_type),

    internal_type =
      blank_to_na(internal_type),

    appt_status =
      blank_to_na(appt_status),

    cancellation_reason =
      blank_to_na(cancellation_reason),

    visit_date =
      parse_excel_date(
        visit_date_raw
      ),

    cancellation_date =
      parse_excel_date(
        cancellation_date_raw
      ),

    appointment_time_seconds =
      parse_time_seconds(
        appointment_time_raw
      ),

    appointment_time =
      format_time_12_hour(
        appointment_time_seconds
      ),

    visit_month =
      lubridate::floor_date(
        visit_date,
        unit = "month"
      ),

    visit_month_label =
      dplyr::if_else(
        is.na(visit_month),
        "Missing visit date",
        format(
          visit_month,
          "%B %Y"
        )
      ),

    appt_department_id =
      extract_bracket_id(
        appt_department
      ),

    appt_department_name =
      remove_bracket_id(
        appt_department
      ),

    clinic_name =
      dplyr::coalesce(
        appt_department_name,
        "Missing or Unrecognized Clinic"
      ),

    clinic_id =
      dplyr::coalesce(
        appt_department_id,
        "MISSING_CLINIC_ID"
      ),

    clinic_key =
      dplyr::coalesce(
        create_grouping_key(
          appt_department_id,
          appt_department_name,
          "CLINIC"
        ),
        "CLINIC_ID_MISSING_CLINIC_ID"
      ),

    external_visit_type_id =
      extract_bracket_id(
        visit_type
      ),

    internal_visit_type_id =
      extract_bracket_id(
        internal_type
      ),

    external_visit_type_name =
      remove_bracket_id(
        visit_type
      ),

    internal_visit_type_name =
      remove_bracket_id(
        internal_type
      ),

    visit_type_id =
      dplyr::coalesce(
        external_visit_type_id,
        internal_visit_type_id
      ),

    visit_type_id_mismatch =
      dplyr::case_when(
        is.na(external_visit_type_id) |
          is.na(internal_visit_type_id) ~
          FALSE,

        external_visit_type_id !=
          internal_visit_type_id ~
          TRUE,

        TRUE ~
          FALSE
      ),

    original_provider_grouping_key =
      dplyr::if_else(
        is.na(provider_resource),
        NA_character_,
        paste0(
          "PROVIDER_COMBINATION_",
          stringr::str_to_upper(
            provider_resource
          )
        )
      ),

    visit_type_grouping_key =
      create_grouping_key(
        visit_type_id,
        internal_visit_type_name,
        "VISIT_TYPE"
      )
  )


# ==============================================================================
# 10. IDENTIFY POSSIBLE DUPLICATE APPOINTMENTS
# ==============================================================================
#
# Flags situations where the same MRN appears more than once for the same date,
# time, original provider combination, and visit type.
#
# Records are retained and flagged rather than automatically removed.
#
# ==============================================================================

visits <- visits |>
  dplyr::group_by(
    clinic_key,
    visit_date,
    appointment_time_seconds,
    original_provider_grouping_key,
    visit_type_grouping_key,
    mrn
  ) |>
  dplyr::mutate(
    matching_rows_for_mrn =
      dplyr::n(),

    possible_duplicate_appointment =
      !is.na(mrn) &
      matching_rows_for_mrn > 1
  ) |>
  dplyr::ungroup()


# ==============================================================================
# 11. INFER GROUP VISITS AND GROUP SESSIONS
# ==============================================================================
#
# A visit is classified as an inferred group visit when at least two distinct
# MRNs share:
#
#   - The same visit date
#   - The same appointment start time
#   - The same original Provider/Resource value
#   - The same visit type
#
# Group inference occurs before joint providers are split into individual
# provider attributions.
#
# ==============================================================================

visits <- visits |>
  dplyr::group_by(
    clinic_key,
    visit_date,
    appointment_time_seconds,
    original_provider_grouping_key,
    visit_type_grouping_key
  ) |>
  dplyr::mutate(
    distinct_mrns_in_session =
      dplyr::n_distinct(
        mrn,
        na.rm = TRUE
      ),

    group_rule_fields_complete =
      !is.na(visit_date) &
      !is.na(appointment_time_seconds) &
      !is.na(original_provider_grouping_key) &
      !is.na(visit_type_grouping_key),

    potential_group_visit =
      group_rule_fields_complete &
      distinct_mrns_in_session >= 2
  ) |>
  dplyr::ungroup() |>
  dplyr::mutate(
    session_id =
      dplyr::if_else(
        group_rule_fields_complete,
        paste(
          clinic_key,
          format(
            visit_date,
            "%Y%m%d"
          ),
          sprintf(
            "%05d",
            as.integer(
              appointment_time_seconds
            )
          ),
          original_provider_grouping_key,
          visit_type_grouping_key,
          sep = "_"
        ),
        NA_character_
      ),

    group_visit_classification =
      dplyr::case_when(
        potential_group_visit ~
          "Inferred group visit",

        !group_rule_fields_complete ~
          "Unclassified/Review",

        TRUE ~
          "Individual visit"
      )
  )


# ==============================================================================
# 12. ASSIGN TIME CATEGORIES
# ==============================================================================
#
# Creates the detailed time categories and identifies appointments included in
# the original 4:00 PM-or-later analysis.
#
# ==============================================================================

four_pm_seconds <- 16 * 3600

four_fifteen_pm_seconds <-
  16 * 3600 +
  15 * 60

visits <- visits |>
  dplyr::mutate(
    time_category =
      dplyr::case_when(
        is.na(appointment_time_seconds) ~
          "Missing or invalid appointment time",

        appointment_time_seconds >= four_pm_seconds &
          appointment_time_seconds < four_fifteen_pm_seconds ~
          "4:00 PM to 4:14 PM",

        appointment_time_seconds >=
          four_fifteen_pm_seconds ~
          "4:15 PM or later",

        TRUE ~
          "Before 4:00 PM"
      ),

    overview_qualifying_visit =
      !is.na(appointment_time_seconds) &
      appointment_time_seconds >=
      four_pm_seconds
  )


# ==============================================================================
# 13. ASSIGN MANAGER STATUS CATEGORIES
# ==============================================================================
#
# Maps the original CPC appointment statuses into three manager categories:
#
#   Comp                 -> Completed
#   Can                  -> Cancelled/No-Show
#   No Show              -> Cancelled/No-Show
#   Left                 -> Other
#   Any unrecognized value -> Other
#
# ==============================================================================

visits <- visits |>
  dplyr::mutate(
    appt_status_standardized =
      stringr::str_to_lower(
        stringr::str_squish(
          dplyr::coalesce(
            appt_status,
            ""
          )
        )
      ),

    manager_status_category =
      dplyr::case_when(
        appt_status_standardized %in% c(
          "comp",
          "complete",
          "completed"
        ) ~
          "Completed",

        appt_status_standardized %in% c(
          "can",
          "cancelled",
          "canceled",
          "no show",
          "no-show",
          "noshow"
        ) ~
          "Cancelled/No-Show",

        stringr::str_detect(
          appt_status_standardized,
          "cancel|no\\s*[-]?\\s*show|noshow"
        ) ~
          "Cancelled/No-Show",

        TRUE ~
          "Other"
      )
  )


# ==============================================================================
# 14. ASSIGN DATA-QUALITY FLAGS
# ==============================================================================
#
# Identifies the first major quality issue associated with each visit record.
# Visits remain in the analysis whenever possible.
#
# ==============================================================================

visits <- visits |>
  dplyr::mutate(
    data_quality_issue =
      dplyr::case_when(
        is.na(visit_date) ~
          "Missing or invalid visit date",

        is.na(appointment_time_seconds) ~
          "Missing or invalid appointment time",

        is.na(provider_resource) ~
          "Missing provider/resource",

        is.na(visit_type_grouping_key) ~
          "Missing visit type",

        is.na(mrn) ~
          "Missing MRN",

        visit_type_id_mismatch ~
          "Visit Type and Type IDs do not match",

        possible_duplicate_appointment ~
          "Possible duplicate appointment",

        TRUE ~
          NA_character_
      )
  )


# ==============================================================================
# 15. ASSIGN GROUP SESSION OUTCOMES
# ==============================================================================
#
# A Completed Group Session has at least one completed patient visit.
#
# A Cancelled/No-Show Group Session has:
#   - No completed patient visits
#   - Every patient visit classified as Cancelled/No-Show
#
# A group is not classified as cancelled solely because one participant
# cancelled or did not attend.
#
# ==============================================================================

visits <- visits |>
  dplyr::group_by(
    session_id
  ) |>
  dplyr::mutate(
    group_session_completed =
      potential_group_visit &
      any(
        manager_status_category ==
          "Completed",
        na.rm = TRUE
      ),

    group_session_cancelled_no_show =
      potential_group_visit &
      !any(
        manager_status_category ==
          "Completed",
        na.rm = TRUE
      ) &
      all(
        manager_status_category ==
          "Cancelled/No-Show"
      ),

    group_session_other_review =
      potential_group_visit &
      !group_session_completed &
      !group_session_cancelled_no_show,

    group_session_outcome =
      dplyr::case_when(
        !potential_group_visit ~
          NA_character_,

        group_session_completed ~
          "Completed Group Session",

        group_session_cancelled_no_show ~
          "Cancelled/No-Show Group Session",

        TRUE ~
          "Other/Review Group Session"
      )
  ) |>
  dplyr::ungroup()


# ==============================================================================
# 16. CREATE THE THREE OVERVIEW TIME-PERIOD DATASETS
# ==============================================================================
#
# Creates separate underlying datasets for:
#
#   1. Before 4:00 PM
#   2. 4:00 PM or Later
#   3. All Time Frames
#
# All Time Frames includes every visit with a valid appointment time.
#
# The existing detailed worksheets continue to focus on 4:00 PM or later.
#
# ==============================================================================

visits_before_4_pm <- visits |>
  dplyr::filter(
    !is.na(appointment_time_seconds),
    appointment_time_seconds <
      four_pm_seconds
  )

visits_4_pm_or_later <- visits |>
  dplyr::filter(
    !is.na(appointment_time_seconds),
    appointment_time_seconds >=
      four_pm_seconds
  )

visits_all_time_frames <- visits |>
  dplyr::filter(
    !is.na(appointment_time_seconds)
  )

qualifying_visits <-
  visits_4_pm_or_later


# ==============================================================================
# 17. CREATE INDIVIDUAL PROVIDER ATTRIBUTIONS
# ==============================================================================
#
# Creates a separate provider-attribution dataset.
#
# Full Clinic continues to count one row per original appointment.
#
# A joint visit involving two providers is attributed once to each provider,
# while still counting as only one Full Clinic visit.
#
# ==============================================================================

create_provider_attributions <- function(
    data
) {

  if (nrow(data) == 0) {
    return(data)
  }

  attribution_data <- purrr::map_dfr(
    seq_len(
      nrow(data)
    ),
    function(row_number_value) {

      visit_row <- data[
        row_number_value,
        ,
        drop = FALSE
      ]

      original_provider_value <-
        visit_row$provider_resource[[1]]

      provider_entries <-
        extract_provider_entries(
          original_provider_value
        )

      purrr::map_dfr(
        provider_entries,
        function(provider_entry) {

          provider_id <-
            extract_bracket_id(
              provider_entry
            )

          provider_name <-
            remove_bracket_id(
              provider_entry
            )

          visit_row |>
            dplyr::mutate(
              original_provider_resource =
                original_provider_value,

              attributed_provider_resource =
                provider_entry,

              attributed_provider_name =
                provider_name,

              attributed_provider_id =
                provider_id,

              attributed_provider_key =
                create_grouping_key(
                  provider_id,
                  provider_name,
                  "PROVIDER"
                ),

              provider_parse_method =
                dplyr::case_when(
                  is.na(original_provider_value) ~
                    "Missing provider/resource",

                  !is.na(provider_id) ~
                    "Parsed from bracketed provider ID",

                  TRUE ~
                    paste0(
                      "Original value retained because ",
                      "no provider ID was found"
                    )
                )
            )
        }
      )
    }
  )

  attribution_data |>
    dplyr::distinct(
      visit_record_id,
      attributed_provider_key,
      .keep_all = TRUE
    )
}

provider_attributions_before_4_pm <-
  create_provider_attributions(
    visits_before_4_pm
  )

provider_attributions_4_pm_or_later <-
  create_provider_attributions(
    visits_4_pm_or_later
  )

provider_attributions_all_time_frames <-
  create_provider_attributions(
    visits_all_time_frames
  )

provider_attributions <-
  provider_attributions_4_pm_or_later


# ==============================================================================
# 18. DEFINE THE OVERVIEW SUMMARY FUNCTION
# ==============================================================================
#
# Calculates the Overview measures for either:
#
#   - Full Clinic unique appointment rows
#   - Individual provider-attribution rows
#
# The function calculates individual visit outcomes, patient-level group visit
# outcomes, and overall group-session outcomes.
#
# ==============================================================================

summarize_overview_rows <- function(
    data,
    reporting_period_value,
    provider_level,
    time_frame_value
) {

  if (provider_level == "clinic") {

    summary_data <- data |>
      dplyr::summarise(
        total_visits =
          dplyr::n_distinct(
            visit_record_id
          ),

        individual_visits =
          safe_distinct_count(
            visit_record_id,
            group_visit_classification ==
              "Individual visit"
          ),

        completed_individual_visits =
          safe_distinct_count(
            visit_record_id,
            group_visit_classification ==
              "Individual visit" &
              manager_status_category ==
              "Completed"
          ),

        cancelled_individual_visits =
          safe_distinct_count(
            visit_record_id,
            group_visit_classification ==
              "Individual visit" &
              manager_status_category ==
              "Cancelled/No-Show"
          ),

        other_individual_visits =
          safe_distinct_count(
            visit_record_id,
            group_visit_classification ==
              "Individual visit" &
              manager_status_category ==
              "Other"
          ),

        inferred_group_visits =
          safe_distinct_count(
            visit_record_id,
            potential_group_visit
          ),

        completed_group_visits =
          safe_distinct_count(
            visit_record_id,
            potential_group_visit &
              manager_status_category ==
              "Completed"
          ),

        cancelled_group_visits =
          safe_distinct_count(
            visit_record_id,
            potential_group_visit &
              manager_status_category ==
              "Cancelled/No-Show"
          ),

        inferred_group_sessions =
          dplyr::n_distinct(
            session_id[
              potential_group_visit
            ],
            na.rm = TRUE
          ),

        completed_group_sessions =
          dplyr::n_distinct(
            session_id[
              potential_group_visit &
                group_session_completed
            ],
            na.rm = TRUE
          ),

        cancelled_group_sessions =
          dplyr::n_distinct(
            session_id[
              potential_group_visit &
                group_session_cancelled_no_show
            ],
            na.rm = TRUE
          ),

        other_review_group_sessions =
          dplyr::n_distinct(
            session_id[
              potential_group_visit &
                group_session_other_review
            ],
            na.rm = TRUE
          ),

        unclassified_review_visits =
          safe_distinct_count(
            visit_record_id,
            group_visit_classification ==
              "Unclassified/Review"
          )
      ) |>
      dplyr::mutate(
        provider_resource =
          "Full Clinic",

        provider_resource_id =
          "Full Clinic",

        provider_sort_order =
          0
      )

  } else {

    summary_data <- data |>
      dplyr::mutate(
        provider_display =
          dplyr::coalesce(
            attributed_provider_name,
            attributed_provider_resource,
            "Missing provider/resource"
          ),

        provider_id_display =
          dplyr::coalesce(
            attributed_provider_id,
            "No bracketed ID"
          ),

        provider_key_display =
          dplyr::coalesce(
            attributed_provider_key,
            paste0(
              "MISSING_PROVIDER_",
              visit_record_id
            )
          )
      ) |>
      dplyr::group_by(
        provider_key_display,
        provider_display,
        provider_id_display
      ) |>
      dplyr::summarise(
        total_visits =
          dplyr::n_distinct(
            visit_record_id
          ),

        individual_visits =
          safe_distinct_count(
            visit_record_id,
            group_visit_classification ==
              "Individual visit"
          ),

        completed_individual_visits =
          safe_distinct_count(
            visit_record_id,
            group_visit_classification ==
              "Individual visit" &
              manager_status_category ==
              "Completed"
          ),

        cancelled_individual_visits =
          safe_distinct_count(
            visit_record_id,
            group_visit_classification ==
              "Individual visit" &
              manager_status_category ==
              "Cancelled/No-Show"
          ),

        other_individual_visits =
          safe_distinct_count(
            visit_record_id,
            group_visit_classification ==
              "Individual visit" &
              manager_status_category ==
              "Other"
          ),

        inferred_group_visits =
          safe_distinct_count(
            visit_record_id,
            potential_group_visit
          ),

        completed_group_visits =
          safe_distinct_count(
            visit_record_id,
            potential_group_visit &
              manager_status_category ==
              "Completed"
          ),

        cancelled_group_visits =
          safe_distinct_count(
            visit_record_id,
            potential_group_visit &
              manager_status_category ==
              "Cancelled/No-Show"
          ),

        inferred_group_sessions =
          dplyr::n_distinct(
            session_id[
              potential_group_visit
            ],
            na.rm = TRUE
          ),

        completed_group_sessions =
          dplyr::n_distinct(
            session_id[
              potential_group_visit &
                group_session_completed
            ],
            na.rm = TRUE
          ),

        cancelled_group_sessions =
          dplyr::n_distinct(
            session_id[
              potential_group_visit &
                group_session_cancelled_no_show
            ],
            na.rm = TRUE
          ),

        other_review_group_sessions =
          dplyr::n_distinct(
            session_id[
              potential_group_visit &
                group_session_other_review
            ],
            na.rm = TRUE
          ),

        unclassified_review_visits =
          safe_distinct_count(
            visit_record_id,
            group_visit_classification ==
              "Unclassified/Review"
          ),

        .groups =
          "drop"
      ) |>
      dplyr::mutate(
        provider_resource =
          provider_display,

        provider_resource_id =
          provider_id_display,

        provider_sort_order =
          1
      )
  }

  summary_data |>
    dplyr::mutate(
      reporting_period =
        reporting_period_value,

      time_frame =
        time_frame_value,

      individual_completion_rate =
        dplyr::if_else(
          individual_visits > 0,
          completed_individual_visits /
            individual_visits,
          NA_real_
        ),

      individual_cancelled_no_show_rate =
        dplyr::if_else(
          individual_visits > 0,
          cancelled_individual_visits /
            individual_visits,
          NA_real_
        ),

      group_session_completion_rate =
        dplyr::if_else(
          inferred_group_sessions > 0,
          completed_group_sessions /
            inferred_group_sessions,
          NA_real_
        ),

      group_session_cancelled_no_show_rate =
        dplyr::if_else(
          inferred_group_sessions > 0,
          cancelled_group_sessions /
            inferred_group_sessions,
          NA_real_
        )
    ) |>
    dplyr::select(
      reporting_period,
      provider_resource,
      provider_resource_id,
      provider_sort_order,
      time_frame,
      total_visits,
      individual_visits,
      completed_individual_visits,
      cancelled_individual_visits,
      other_individual_visits,
      inferred_group_visits,
      completed_group_visits,
      cancelled_group_visits,
      inferred_group_sessions,
      completed_group_sessions,
      cancelled_group_sessions,
      other_review_group_sessions,
      unclassified_review_visits,
      individual_completion_rate,
      individual_cancelled_no_show_rate,
      group_session_completion_rate,
      group_session_cancelled_no_show_rate
    )
}


# ==============================================================================
# 19. BUILD THE MANAGER OVERVIEW
# ==============================================================================
#
# Creates Overview rows for:
#
#   - Full Clinic
#   - Every individually attributed provider
#   - All Months Combined
#   - Every reporting month
#   - Before 4:00 PM
#   - 4:00 PM or Later
#   - All Time Frames
#
# ==============================================================================

create_overview_for_time_frame <- function(
    clinic_data,
    provider_data,
    time_frame_label,
    time_frame_sort_value
) {

  all_months_clinic <-
    summarize_overview_rows(
      data =
        clinic_data,

      reporting_period_value =
        "All Months Combined",

      provider_level =
        "clinic",

      time_frame_value =
        time_frame_label
    )

  all_months_providers <-
    summarize_overview_rows(
      data =
        provider_data,

      reporting_period_value =
        "All Months Combined",

      provider_level =
        "provider",

      time_frame_value =
        time_frame_label
    )

  monthly_clinic <- clinic_data |>
    dplyr::group_split(
      visit_month_label
    ) |>
    purrr::map_dfr(
      function(month_data) {

        month_label <- unique(
          month_data$visit_month_label
        )[1]

        summarize_overview_rows(
          data =
            month_data,

          reporting_period_value =
            month_label,

          provider_level =
            "clinic",

          time_frame_value =
            time_frame_label
        )
      }
    )

  monthly_providers <- provider_data |>
    dplyr::group_split(
      visit_month_label
    ) |>
    purrr::map_dfr(
      function(month_data) {

        month_label <- unique(
          month_data$visit_month_label
        )[1]

        summarize_overview_rows(
          data =
            month_data,

          reporting_period_value =
            month_label,

          provider_level =
            "provider",

          time_frame_value =
            time_frame_label
        )
      }
    )

  dplyr::bind_rows(
    all_months_clinic,
    all_months_providers,
    monthly_clinic,
    monthly_providers
  ) |>
    dplyr::mutate(
      time_frame_sort =
        time_frame_sort_value
    )
}

overview_before_4_pm <-
  create_overview_for_time_frame(
    clinic_data =
      visits_before_4_pm,

    provider_data =
      provider_attributions_before_4_pm,

    time_frame_label =
      "Before 4:00 PM",

    time_frame_sort_value =
      1
  )

overview_4_pm_or_later <-
  create_overview_for_time_frame(
    clinic_data =
      visits_4_pm_or_later,

    provider_data =
      provider_attributions_4_pm_or_later,

    time_frame_label =
      "4:00 PM or Later",

    time_frame_sort_value =
      2
  )

overview_all_time_frames <-
  create_overview_for_time_frame(
    clinic_data =
      visits_all_time_frames,

    provider_data =
      provider_attributions_all_time_frames,

    time_frame_label =
      "All Time Frames",

    time_frame_sort_value =
      3
  )

overview_data <- dplyr::bind_rows(
  overview_before_4_pm,
  overview_4_pm_or_later,
  overview_all_time_frames
) |>
  dplyr::mutate(
    reporting_period_sort =
      dplyr::case_when(
        reporting_period ==
          "All Months Combined" ~
          as.Date("1900-01-01"),

        reporting_period ==
          "Missing visit date" ~
          as.Date("9999-12-31"),

        TRUE ~
          suppressWarnings(
            as.Date(
              paste0(
                "01 ",
                reporting_period
              ),
              format = "%d %B %Y"
            )
          )
      )
  ) |>
  dplyr::arrange(
    reporting_period_sort,
    provider_sort_order,
    provider_resource,
    time_frame_sort
  ) |>
  dplyr::transmute(
    `Reporting Period` =
      reporting_period,

    `Provider/Resource` =
      provider_resource,

    `Provider/Resource ID` =
      provider_resource_id,

    `Time Frame` =
      time_frame,

    `Total Visits` =
      total_visits,

    `Individual Visits` =
      individual_visits,

    `Completed Individual Visits` =
      completed_individual_visits,

    `Cancelled/No-Show Individual Visits` =
      cancelled_individual_visits,

    `Other Individual Visits` =
      other_individual_visits,

    `Inferred Group Visits` =
      inferred_group_visits,

    `Completed Group Visits` =
      completed_group_visits,

    `Cancelled/No-Show Group Visits` =
      cancelled_group_visits,

    `Inferred Group Sessions` =
      inferred_group_sessions,

    `Completed Group Sessions` =
      completed_group_sessions,

    `Cancelled/No-Show Group Sessions` =
      cancelled_group_sessions,

    `Other/Review Group Sessions` =
      other_review_group_sessions,

    `Unclassified/Review Visits` =
      unclassified_review_visits,

    `Individual Completion Rate` =
      individual_completion_rate,

    `Individual Cancelled/No-Show Rate` =
      individual_cancelled_no_show_rate,

    `Group Session Completion Rate` =
      group_session_completion_rate,

    `Group Session Cancelled/No-Show Rate` =
      group_session_cancelled_no_show_rate
  )


# ==============================================================================
# 20. VALIDATE OVERVIEW RECONCILIATION
# ==============================================================================
#
# Confirms that:
#
#   Before 4:00 PM + 4:00 PM or Later = All Time Frames
#
# Also confirms that completed, cancelled/no-show, and other/review group
# sessions reconcile to the total inferred group-session count.
#
# ==============================================================================

overview_reconciliation <- overview_data |>
  dplyr::select(
    `Reporting Period`,
    `Provider/Resource`,
    `Provider/Resource ID`,
    `Time Frame`,
    `Total Visits`
  ) |>
  tidyr::pivot_wider(
    names_from =
      `Time Frame`,

    values_from =
      `Total Visits`,

    values_fill =
      0
  ) |>
  dplyr::mutate(
    expected_all_time_frames =
      `Before 4:00 PM` +
      `4:00 PM or Later`,

    time_frame_difference =
      `All Time Frames` -
      expected_all_time_frames
  )

time_frame_reconciliation_errors <-
  overview_reconciliation |>
  dplyr::filter(
    time_frame_difference != 0
  )

if (
  nrow(
    time_frame_reconciliation_errors
  ) > 0
) {

  stop(
    paste0(
      "Overview time frames did not reconcile. ",
      "Before 4:00 PM plus 4:00 PM or Later must equal ",
      "All Time Frames for every reporting period and provider."
    ),
    call. = FALSE
  )
}

group_session_reconciliation_errors <-
  overview_data |>
  dplyr::filter(
    `Completed Group Sessions` +
      `Cancelled/No-Show Group Sessions` +
      `Other/Review Group Sessions` !=
      `Inferred Group Sessions`
  )

if (
  nrow(
    group_session_reconciliation_errors
  ) > 0
) {

  stop(
    paste0(
      "Group-session outcomes did not reconcile. ",
      "Completed, Cancelled/No-Show, and Other/Review ",
      "Group Sessions must equal Inferred Group Sessions."
    ),
    call. = FALSE
  )
}


# ==============================================================================
# 21. CREATE THE MONTHLY LATE-VISIT COMPARISON
# ==============================================================================
#
# Creates a clinic-level monthly summary focused on the original scheduling
# question: visits beginning at 4:00 PM or later.
#
# ==============================================================================

monthly_comparison <- qualifying_visits |>
  dplyr::group_by(
    visit_month,
    visit_month_label
  ) |>
  dplyr::summarise(
    total_visits =
      dplyr::n_distinct(
        visit_record_id
      ),

    visits_exactly_4_pm =
      safe_distinct_count(
        visit_record_id,
        time_category ==
          "4:00 PM to 4:14 PM"
      ),

    visits_4_15_or_later =
      safe_distinct_count(
        visit_record_id,
        time_category ==
          "4:15 PM or later"
      ),

    individual_visits =
      safe_distinct_count(
        visit_record_id,
        group_visit_classification ==
          "Individual visit"
      ),

    completed_individual_visits =
      safe_distinct_count(
        visit_record_id,
        group_visit_classification ==
          "Individual visit" &
          manager_status_category ==
          "Completed"
      ),

    cancelled_individual_visits =
      safe_distinct_count(
        visit_record_id,
        group_visit_classification ==
          "Individual visit" &
          manager_status_category ==
          "Cancelled/No-Show"
      ),

    inferred_group_visits =
      safe_distinct_count(
        visit_record_id,
        potential_group_visit
      ),

    inferred_group_sessions =
      dplyr::n_distinct(
        session_id[
          potential_group_visit
        ],
        na.rm = TRUE
      ),

    completed_group_sessions =
      dplyr::n_distinct(
        session_id[
          potential_group_visit &
            group_session_completed
        ],
        na.rm = TRUE
      ),

    cancelled_group_sessions =
      dplyr::n_distinct(
        session_id[
          potential_group_visit &
            group_session_cancelled_no_show
        ],
        na.rm = TRUE
      ),

    .groups =
      "drop"
  ) |>
  dplyr::arrange(
    visit_month
  ) |>
  dplyr::transmute(
    `Reporting Period` =
      visit_month_label,

    `Total Visits at 4:00 PM or Later` =
      total_visits,

    `Exactly 4:00 PM` =
      visits_exactly_4_pm,

    `4:15 PM or Later` =
      visits_4_15_or_later,

    `Individual Visits` =
      individual_visits,

    `Completed Individual Visits` =
      completed_individual_visits,

    `Cancelled/No-Show Individual Visits` =
      cancelled_individual_visits,

    `Inferred Group Visits` =
      inferred_group_visits,

    `Inferred Group Sessions` =
      inferred_group_sessions,

    `Completed Group Sessions` =
      completed_group_sessions,

    `Cancelled/No-Show Group Sessions` =
      cancelled_group_sessions
  )


# ==============================================================================
# 22. CREATE PROVIDER DETAIL
# ==============================================================================
#
# Creates detailed provider-attribution results for visits beginning at 4:00 PM
# or later, broken down by reporting month, time category, and visit type.
#
# ==============================================================================

provider_detail <- provider_attributions |>
  dplyr::mutate(
    provider_display =
      dplyr::coalesce(
        attributed_provider_name,
        attributed_provider_resource,
        "Missing provider/resource"
      ),

    provider_id_display =
      dplyr::coalesce(
        attributed_provider_id,
        "No bracketed ID"
      )
  ) |>
  dplyr::group_by(
    visit_month,
    visit_month_label,
    provider_display,
    provider_id_display,
    time_category,
    external_visit_type_name,
    internal_visit_type_name,
    visit_type_id
  ) |>
  dplyr::summarise(
    attributed_visits =
      dplyr::n_distinct(
        visit_record_id
      ),

    individual_visits =
      safe_distinct_count(
        visit_record_id,
        group_visit_classification ==
          "Individual visit"
      ),

    inferred_group_visits =
      safe_distinct_count(
        visit_record_id,
        potential_group_visit
      ),

    inferred_group_sessions =
      dplyr::n_distinct(
        session_id[
          potential_group_visit
        ],
        na.rm = TRUE
      ),

    completed_visits =
      safe_distinct_count(
        visit_record_id,
        manager_status_category ==
          "Completed"
      ),

    cancelled_no_show_visits =
      safe_distinct_count(
        visit_record_id,
        manager_status_category ==
          "Cancelled/No-Show"
      ),

    other_status_visits =
      safe_distinct_count(
        visit_record_id,
        manager_status_category ==
          "Other"
      ),

    .groups =
      "drop"
  ) |>
  dplyr::arrange(
    visit_month,
    provider_display,
    time_category,
    dplyr::desc(
      attributed_visits
    )
  ) |>
  dplyr::transmute(
    `Reporting Period` =
      visit_month_label,

    `Provider/Resource` =
      provider_display,

    `Provider/Resource ID` =
      provider_id_display,

    `Time Category` =
      time_category,

    `External Visit Type` =
      external_visit_type_name,

    `Internal Visit Type` =
      internal_visit_type_name,

    `Visit Type ID` =
      visit_type_id,

    `Attributed Visits` =
      attributed_visits,

    `Individual Visits` =
      individual_visits,

    `Inferred Group Visits` =
      inferred_group_visits,

    `Inferred Group Sessions` =
      inferred_group_sessions,

    `Completed Visits` =
      completed_visits,

    `Cancelled/No-Show Visits` =
      cancelled_no_show_visits,

    `Other Status Visits` =
      other_status_visits
  )


# ==============================================================================
# 23. CREATE STATUS DETAIL
# ==============================================================================
#
# Creates status-level results for:
#
#   - All Months Combined
#   - Each individual reporting month
#
# The Inferred Group Sessions column shows how many sessions included at least
# one patient visit with the displayed status. A session can appear under more
# than one status and this column should not be summed across status rows.
#
# ==============================================================================

create_status_detail <- function(
    data,
    reporting_period_value
) {

  data |>
    dplyr::mutate(
      original_status_display =
        dplyr::coalesce(
          appt_status,
          "Missing status"
        )
    ) |>
    dplyr::group_by(
      original_status_display,
      manager_status_category
    ) |>
    dplyr::summarise(
      visits =
        dplyr::n_distinct(
          visit_record_id
        ),

      individual_visits =
        safe_distinct_count(
          visit_record_id,
          group_visit_classification ==
            "Individual visit"
        ),

      inferred_group_visits =
        safe_distinct_count(
          visit_record_id,
          potential_group_visit
        ),

      inferred_group_sessions =
        dplyr::n_distinct(
          session_id[
            potential_group_visit &
              !is.na(session_id)
          ],
          na.rm = TRUE
        ),

      unclassified_review_visits =
        safe_distinct_count(
          visit_record_id,
          group_visit_classification ==
            "Unclassified/Review"
        ),

      .groups =
        "drop"
    ) |>
    dplyr::mutate(
      reporting_period =
        reporting_period_value
    )
}

status_detail_all_months <-
  create_status_detail(
    data =
      qualifying_visits,

    reporting_period_value =
      "All Months Combined"
  )

status_detail_monthly <- qualifying_visits |>
  dplyr::group_split(
    visit_month_label
  ) |>
  purrr::map_dfr(
    function(month_data) {

      month_label <- unique(
        month_data$visit_month_label
      )[1]

      create_status_detail(
        data =
          month_data,

        reporting_period_value =
          month_label
      )
    }
  )

status_detail <- dplyr::bind_rows(
  status_detail_all_months,
  status_detail_monthly
) |>
  dplyr::mutate(
    reporting_period_sort =
      dplyr::case_when(
        reporting_period ==
          "All Months Combined" ~
          as.Date("1900-01-01"),

        reporting_period ==
          "Missing visit date" ~
          as.Date("9999-12-31"),

        TRUE ~
          suppressWarnings(
            as.Date(
              paste0(
                "01 ",
                reporting_period
              ),
              format = "%d %B %Y"
            )
          )
      ),

    status_category_sort =
      dplyr::case_when(
        manager_status_category ==
          "Completed" ~
          1,

        manager_status_category ==
          "Cancelled/No-Show" ~
          2,

        manager_status_category ==
          "Other" ~
          3,

        TRUE ~
          4
      )
  ) |>
  dplyr::arrange(
    reporting_period_sort,
    status_category_sort,
    dplyr::desc(visits),
    original_status_display
  ) |>
  dplyr::transmute(
    `Reporting Period` =
      reporting_period,

    `Original Appointment Status` =
      original_status_display,

    `Manager Status Category` =
      manager_status_category,

    `Visits` =
      visits,

    `Individual Visits` =
      individual_visits,

    `Inferred Group Visits` =
      inferred_group_visits,

    `Inferred Group Sessions` =
      inferred_group_sessions,

    `Unclassified/Review Visits` =
      unclassified_review_visits
  )


# ==============================================================================
# 24. CREATE GROUP SESSION DETAIL
# ==============================================================================
#
# Creates one row per inferred group session beginning at 4:00 PM or later.
# Includes patient-level status counts and the overall group-session outcome.
#
# ==============================================================================

group_sessions <- qualifying_visits |>
  dplyr::filter(
    potential_group_visit
  ) |>
  dplyr::group_by(
    session_id,
    visit_date,
    appointment_time,
    appointment_time_seconds,
    provider_resource,
    external_visit_type_name,
    internal_visit_type_name,
    visit_type_id,
    time_category
  ) |>
  dplyr::summarise(
    scheduled_visit_records =
      dplyr::n_distinct(
        visit_record_id
      ),

    distinct_patients =
      dplyr::n_distinct(
        mrn,
        na.rm = TRUE
      ),

    group_session_outcome =
      dplyr::first(
        group_session_outcome
      ),

    completed_visits =
      safe_distinct_count(
        visit_record_id,
        manager_status_category ==
          "Completed"
      ),

    cancelled_no_show_visits =
      safe_distinct_count(
        visit_record_id,
        manager_status_category ==
          "Cancelled/No-Show"
      ),

    other_status_visits =
      safe_distinct_count(
        visit_record_id,
        manager_status_category ==
          "Other"
      ),

    appointment_statuses =
      paste(
        sort(
          unique(
            stats::na.omit(
              appt_status
            )
          )
        ),
        collapse = "; "
      ),

    source_files =
      paste(
        sort(
          unique(
            source_file
          )
        ),
        collapse = "; "
      ),

    possible_duplicate_records =
      sum(
        possible_duplicate_appointment,
        na.rm = TRUE
      ),

    .groups =
      "drop"
  ) |>
  dplyr::arrange(
    visit_date,
    appointment_time_seconds,
    provider_resource
  ) |>
  dplyr::transmute(
    `Session ID` =
      session_id,

    `Visit Date` =
      visit_date,

    `Start Time` =
      appointment_time,

    `Time Category` =
      time_category,

    `Original Provider/Resource` =
      provider_resource,

    `External Visit Type` =
      external_visit_type_name,

    `Internal Visit Type` =
      internal_visit_type_name,

    `Visit Type ID` =
      visit_type_id,

    `Scheduled Visit Records` =
      scheduled_visit_records,

    `Distinct Patients` =
      distinct_patients,

    `Group Session Outcome` =
      group_session_outcome,

    `Completed Visits` =
      completed_visits,

    `Cancelled/No-Show Visits` =
      cancelled_no_show_visits,

    `Other Status Visits` =
      other_status_visits,

    `Appointment Statuses` =
      appointment_statuses,

    `Possible Duplicate Records` =
      possible_duplicate_records,

    `Source Files` =
      source_files
  )


# ==============================================================================
# 25. CREATE THE PROVIDER ATTRIBUTION AUDIT
# ==============================================================================
#
# Shows how every original joint or single Provider/Resource value was assigned
# to an individual provider for the late-visit provider summaries.
#
# ==============================================================================

provider_attribution_audit <- provider_attributions |>
  dplyr::arrange(
    visit_date,
    appointment_time_seconds,
    original_provider_resource,
    attributed_provider_name
  ) |>
  dplyr::transmute(
    `Visit Record ID` =
      visit_record_id,

    `Source File` =
      source_file,

    `Source Row` =
      source_row,

    `Visit Date` =
      visit_date,

    `Start Time` =
      appointment_time,

    `Original Provider/Resource` =
      original_provider_resource,

    `Attributed Provider/Resource` =
      attributed_provider_resource,

    `Attributed Provider Name` =
      attributed_provider_name,

    `Attributed Provider ID` =
      attributed_provider_id,

    `Provider Parse Method` =
      provider_parse_method,

    `Group Classification` =
      group_visit_classification,

    `Original Appointment Status` =
      appt_status,

    `Manager Status Category` =
      manager_status_category
  )


# ==============================================================================
# 26. CREATE OPERATIONAL RAW DATA (DIRECT PATIENT IDENTIFIERS OMITTED)
# ==============================================================================
#
# Creates a visit-level worksheet containing one row per original appointment.
#
# MRN and Patient are intentionally excluded. All other original fields and the
# most useful derived fields are retained for manager drill-down.
#
# ==============================================================================

raw_data <- visits |>
  dplyr::arrange(
    visit_date,
    appointment_time_seconds,
    provider_resource
  ) |>
  dplyr::transmute(
    `Visit Record ID` =
      visit_record_id,

    `Source File` =
      source_file,

    `Source Row` =
      source_row,

    `Appt Department` =
      appt_department,

    `Provider Department` =
      provider_department,

    `Clinic` =
      clinic_name,

    `Clinic ID` =
      clinic_id,

    `Provider/Resource` =
      provider_resource,

    `Visit Type` =
      visit_type,

    `Type` =
      internal_type,

    `Visit Date` =
      visit_date,

    `Time` =
      appointment_time,

    `Appt Status` =
      appt_status,

    `Canc Date` =
      cancellation_date,

    `Canc Reason` =
      cancellation_reason,

    `Visit Month` =
      visit_month_label,

    `Time Category` =
      time_category,

    `4:00 PM or Later` =
      overview_qualifying_visit,

    `Manager Status Category` =
      manager_status_category,

    `Group Classification` =
      group_visit_classification,

    `Distinct Patients in Session` =
      distinct_mrns_in_session,

    `Session ID` =
      session_id,

    `Group Session Outcome` =
      group_session_outcome,

    `Possible Duplicate Appointment` =
      possible_duplicate_appointment,

    `Visit Type ID Mismatch` =
      visit_type_id_mismatch,

    `Data Quality Issue` =
      data_quality_issue
  )


# ==============================================================================
# END OF VERSION 3 - BASE ANALYSIS SECTION
# Paste Section 3 immediately below this line.
# ==============================================================================

# ==============================================================================
# VERSION 3 - REPORT OUTPUT SECTION
# ==============================================================================



# ==============================================================================
# 26A. VERSION 3 PORTFOLIO, TIME-FRAME, PATIENT, CANCELLATION, AND CHART LOGIC
# ==============================================================================
#
# This section extends the stable Version 2 appointment logic. It calculates
# seven overlapping time frames independently, adds portfolio and clinic
# reporting levels, calculates privacy-safe unique-patient measures, summarizes
# explicit cancellations, and prepares hourly service-activity chart data.
#
# ==============================================================================

portfolio_display_name <- "Full Portfolio"
all_providers_display_name <- "All Providers"
missing_clinic_display_name <- "Missing or Unrecognized Clinic"

calculate_start_hour <- function(seconds_after_midnight) {
  result <- rep(NA_integer_, length(seconds_after_midnight))
  valid_rows <- !is.na(seconds_after_midnight)
  result[valid_rows] <- as.integer(
    floor(seconds_after_midnight[valid_rows] / 3600)
  )
  result
}

format_hour_label <- function(hour_value) {
  result <- rep(NA_character_, length(hour_value))
  valid_rows <- !is.na(hour_value)
  if (!any(valid_rows)) return(result)
  valid_hours <- hour_value[valid_rows]
  hour_12 <- valid_hours %% 12
  hour_12[hour_12 == 0] <- 12
  result[valid_rows] <- paste0(
    hour_12,
    ":00 ",
    ifelse(valid_hours < 12, "AM", "PM")
  )
  result
}

safe_patient_count <- function(patient_ids, condition = NULL) {
  if (is.null(condition)) {
    selected_ids <- patient_ids[!is.na(patient_ids)]
  } else {
    selected_ids <- patient_ids[
      !is.na(patient_ids) & !is.na(condition) & condition
    ]
  }
  dplyr::n_distinct(selected_ids, na.rm = TRUE)
}

visits <- visits |>
  dplyr::mutate(
    start_hour = calculate_start_hour(appointment_time_seconds),
    start_hour_label = format_hour_label(start_hour),
    explicitly_cancelled_visit =
      appt_status_standardized %in% c("can", "cancelled", "canceled") |
      stringr::str_detect(appt_status_standardized, "^cancel"),
    no_show_visit =
      appt_status_standardized %in% c("no show", "no-show", "noshow"),
    cancellation_reason_display = dplyr::case_when(
      explicitly_cancelled_visit & is.na(cancellation_reason) ~
        "Missing cancellation reason",
      explicitly_cancelled_visit ~ cancellation_reason,
      TRUE ~ NA_character_
    )
  )

overview_time_frames <- tibble::tribble(
  ~time_frame_order, ~time_frame, ~start_seconds, ~end_seconds,
  1L, "Full Day", NA_real_, NA_real_,
  2L, "8:00 AM to 4:00 PM", 8 * 3600, 16 * 3600,
  3L, "9:00 AM to 5:00 PM", 9 * 3600, 17 * 3600,
  4L, "8:00 AM to 9:00 AM", 8 * 3600, 9 * 3600,
  5L, "Before 4:00 PM", NA_real_, 16 * 3600,
  6L, "4:00 PM or Later", 16 * 3600, NA_real_,
  7L, "4:30 PM or Later", 16 * 3600 + 30 * 60, NA_real_
)

filter_for_time_frame <- function(data, start_seconds, end_seconds) {
  filtered_data <- data |>
    dplyr::filter(!is.na(appointment_time_seconds))
  if (!is.na(start_seconds)) {
    filtered_data <- filtered_data |>
      dplyr::filter(appointment_time_seconds >= start_seconds)
  }
  if (!is.na(end_seconds)) {
    filtered_data <- filtered_data |>
      dplyr::filter(appointment_time_seconds < end_seconds)
  }
  filtered_data
}

visits_by_time_frame <- purrr::pmap_dfr(
  overview_time_frames,
  function(time_frame_order, time_frame, start_seconds, end_seconds) {
    filter_for_time_frame(visits, start_seconds, end_seconds) |>
      dplyr::mutate(
        time_frame_order = time_frame_order,
        time_frame = time_frame
      )
  }
)

provider_attributions_all_visits <- create_provider_attributions(visits)

provider_attributions_by_time_frame <- purrr::pmap_dfr(
  overview_time_frames,
  function(time_frame_order, time_frame, start_seconds, end_seconds) {
    create_provider_attributions(
      filter_for_time_frame(visits, start_seconds, end_seconds)
    ) |>
      dplyr::mutate(
        time_frame_order = time_frame_order,
        time_frame = time_frame
      )
  }
)

add_reporting_periods <- function(data) {
  dplyr::bind_rows(
    data |>
      dplyr::mutate(
        reporting_period = "All Months Combined",
        reporting_period_date = as.Date("1900-01-01"),
        reporting_period_order = 0L
      ),
    data |>
      dplyr::mutate(
        reporting_period = visit_month_label,
        reporting_period_date = visit_month,
        reporting_period_order = 1L
      )
  )
}

summarize_v3_activity <- function(data) {
  data |>
    dplyr::summarise(
      total_visits = dplyr::n_distinct(visit_record_id),
      unique_patients_scheduled = safe_patient_count(mrn),
      unique_patients_served = safe_patient_count(
        mrn,
        manager_status_category == "Completed"
      ),
      completed_patient_level_visits = safe_distinct_count(
        visit_record_id,
        manager_status_category == "Completed"
      ),
      individual_visits = safe_distinct_count(
        visit_record_id,
        group_visit_classification == "Individual visit"
      ),
      completed_individual_visits = safe_distinct_count(
        visit_record_id,
        group_visit_classification == "Individual visit" &
          manager_status_category == "Completed"
      ),
      cancelled_individual_visits = safe_distinct_count(
        visit_record_id,
        group_visit_classification == "Individual visit" &
          manager_status_category == "Cancelled/No-Show"
      ),
      other_individual_visits = safe_distinct_count(
        visit_record_id,
        group_visit_classification == "Individual visit" &
          manager_status_category == "Other"
      ),
      inferred_group_visits = safe_distinct_count(
        visit_record_id,
        potential_group_visit
      ),
      completed_group_visits = safe_distinct_count(
        visit_record_id,
        potential_group_visit & manager_status_category == "Completed"
      ),
      cancelled_group_visits = safe_distinct_count(
        visit_record_id,
        potential_group_visit & manager_status_category == "Cancelled/No-Show"
      ),
      inferred_group_sessions = dplyr::n_distinct(
        session_id[potential_group_visit],
        na.rm = TRUE
      ),
      completed_group_sessions = dplyr::n_distinct(
        session_id[potential_group_visit & group_session_completed],
        na.rm = TRUE
      ),
      cancelled_group_sessions = dplyr::n_distinct(
        session_id[potential_group_visit & group_session_cancelled_no_show],
        na.rm = TRUE
      ),
      other_review_group_sessions = dplyr::n_distinct(
        session_id[potential_group_visit & group_session_other_review],
        na.rm = TRUE
      ),
      unclassified_review_visits = safe_distinct_count(
        visit_record_id,
        group_visit_classification == "Unclassified/Review"
      ),
      .groups = "drop"
    ) |>
    dplyr::mutate(
      completed_visits_per_patient_served = dplyr::if_else(
        unique_patients_served > 0,
        completed_patient_level_visits / unique_patients_served,
        NA_real_
      ),
      individual_completion_rate = dplyr::if_else(
        individual_visits > 0,
        completed_individual_visits / individual_visits,
        NA_real_
      ),
      individual_cancelled_no_show_rate = dplyr::if_else(
        individual_visits > 0,
        cancelled_individual_visits / individual_visits,
        NA_real_
      ),
      group_session_completion_rate = dplyr::if_else(
        inferred_group_sessions > 0,
        completed_group_sessions / inferred_group_sessions,
        NA_real_
      ),
      group_session_cancelled_no_show_rate = dplyr::if_else(
        inferred_group_sessions > 0,
        cancelled_group_sessions / inferred_group_sessions,
        NA_real_
      )
    )
}

v3_visits <- add_reporting_periods(visits_by_time_frame)
v3_providers <- add_reporting_periods(provider_attributions_by_time_frame)

overview_portfolio_v3 <- v3_visits |>
  dplyr::group_by(
    reporting_period,
    reporting_period_date,
    reporting_period_order,
    time_frame_order,
    time_frame
  ) |>
  summarize_v3_activity() |>
  dplyr::mutate(
    reporting_level = "Full Portfolio",
    reporting_level_order = 1L,
    clinic_name = portfolio_display_name,
    clinic_id = "FULL_PORTFOLIO",
    provider_name = all_providers_display_name,
    provider_id = all_providers_display_name
  ) |>
  dplyr::ungroup()

overview_clinic_v3 <- v3_visits |>
  dplyr::group_by(
    reporting_period,
    reporting_period_date,
    reporting_period_order,
    clinic_key,
    clinic_name,
    clinic_id,
    time_frame_order,
    time_frame
  ) |>
  summarize_v3_activity() |>
  dplyr::mutate(
    reporting_level = "Full Clinic",
    reporting_level_order = 2L,
    provider_name = "Full Clinic",
    provider_id = "Full Clinic"
  ) |>
  dplyr::ungroup()

overview_provider_v3 <- v3_providers |>
  dplyr::group_by(
    reporting_period,
    reporting_period_date,
    reporting_period_order,
    clinic_key,
    clinic_name,
    clinic_id,
    attributed_provider_key,
    attributed_provider_name,
    attributed_provider_id,
    time_frame_order,
    time_frame
  ) |>
  summarize_v3_activity() |>
  dplyr::mutate(
    reporting_level = "Provider",
    reporting_level_order = 3L,
    provider_name = attributed_provider_name,
    provider_id = attributed_provider_id
  ) |>
  dplyr::ungroup()

overview_v3_base <- dplyr::bind_rows(
  overview_portfolio_v3,
  overview_clinic_v3,
  overview_provider_v3
)

overview_data <- overview_v3_base |>
  dplyr::arrange(
    reporting_period_order,
    reporting_period_date,
    reporting_level_order,
    clinic_name,
    provider_name,
    time_frame_order
  ) |>
  dplyr::transmute(
    `Reporting Period` = reporting_period,
    `Reporting Level` = reporting_level,
    `Clinic` = clinic_name,
    `Clinic ID` = clinic_id,
    `Provider/Resource` = provider_name,
    `Provider/Resource ID` = provider_id,
    `Time Frame` = time_frame,
    `Total Visits` = total_visits,
    `Unique Patients Scheduled` = unique_patients_scheduled,
    `Unique Patients Served` = unique_patients_served,
    `Completed Patient-Level Visits` = completed_patient_level_visits,
    `Completed Visits per Patient Served` = completed_visits_per_patient_served,
    `Individual Visits` = individual_visits,
    `Completed Individual Visits` = completed_individual_visits,
    `Cancelled/No-Show Individual Visits` = cancelled_individual_visits,
    `Other Individual Visits` = other_individual_visits,
    `Inferred Group Visits` = inferred_group_visits,
    `Completed Group Visits` = completed_group_visits,
    `Cancelled/No-Show Group Visits` = cancelled_group_visits,
    `Inferred Group Sessions` = inferred_group_sessions,
    `Completed Group Sessions` = completed_group_sessions,
    `Cancelled/No-Show Group Sessions` = cancelled_group_sessions,
    `Other/Review Group Sessions` = other_review_group_sessions,
    `Unclassified/Review Visits` = unclassified_review_visits,
    `Individual Completion Rate` = individual_completion_rate,
    `Individual Cancelled/No-Show Rate` = individual_cancelled_no_show_rate,
    `Group Session Completion Rate` = group_session_completion_rate,
    `Group Session Cancelled/No-Show Rate` = group_session_cancelled_no_show_rate
  )

unique_patients_data <- overview_v3_base |>
  dplyr::group_by(
    reporting_period,
    reporting_level,
    clinic_name,
    clinic_id,
    provider_name,
    provider_id
  ) |>
  dplyr::mutate(
    full_day_patients_served = unique_patients_served[
      match("Full Day", time_frame)
    ],
    percentage_of_full_day_patients_served = dplyr::if_else(
      full_day_patients_served > 0,
      unique_patients_served / full_day_patients_served,
      NA_real_
    )
  ) |>
  dplyr::ungroup() |>
  dplyr::arrange(
    reporting_period_order,
    reporting_period_date,
    reporting_level_order,
    clinic_name,
    provider_name,
    time_frame_order
  ) |>
  dplyr::transmute(
    `Reporting Period` = reporting_period,
    `Reporting Level` = reporting_level,
    `Clinic` = clinic_name,
    `Clinic ID` = clinic_id,
    `Provider/Resource` = provider_name,
    `Provider/Resource ID` = provider_id,
    `Time Frame` = time_frame,
    `Total Visits` = total_visits,
    `Unique Patients Scheduled` = unique_patients_scheduled,
    `Unique Patients Served` = unique_patients_served,
    `Completed Patient-Level Visits` = completed_patient_level_visits,
    `Completed Visits per Patient Served` = completed_visits_per_patient_served,
    `Percentage of Full-Day Patients Served` = percentage_of_full_day_patients_served
  )

cancelled_visits <- add_reporting_periods(
  visits |>
    dplyr::filter(explicitly_cancelled_visit)
)

summarize_cancellations_v3 <- function(data) {
  data |>
    dplyr::group_by(
      reporting_period,
      reporting_period_date,
      reporting_period_order,
      cancellation_reason_display
    ) |>
    dplyr::summarise(
      cancelled_visits = dplyr::n_distinct(visit_record_id),
      cancelled_individual_visits = safe_distinct_count(
        visit_record_id,
        group_visit_classification == "Individual visit"
      ),
      cancelled_group_visits = safe_distinct_count(
        visit_record_id,
        potential_group_visit
      ),
      affected_group_sessions = dplyr::n_distinct(
        session_id[potential_group_visit],
        na.rm = TRUE
      ),
      .groups = "drop"
    ) |>
    dplyr::group_by(reporting_period) |>
    dplyr::mutate(
      percentage_of_cancellations = cancelled_visits / sum(cancelled_visits)
    ) |>
    dplyr::ungroup()
}

cancellations_portfolio <- cancelled_visits |>
  summarize_cancellations_v3() |>
  dplyr::mutate(
    reporting_level = "Full Portfolio",
    reporting_level_order = 1L,
    cancellation_clinic = portfolio_display_name,
    cancellation_clinic_id = "FULL_PORTFOLIO",
    cancellation_provider = all_providers_display_name,
    cancellation_provider_id = all_providers_display_name
  )

cancellations_clinic <- cancelled_visits |>
  dplyr::group_split(clinic_key) |>
  purrr::map_dfr(
    function(clinic_data) {
      clinic_data |>
        summarize_cancellations_v3() |>
        dplyr::mutate(
          reporting_level = "Full Clinic",
          reporting_level_order = 2L,
          cancellation_clinic = unique(clinic_data$clinic_name)[1],
          cancellation_clinic_id = unique(clinic_data$clinic_id)[1],
          cancellation_provider = "Full Clinic",
          cancellation_provider_id = "Full Clinic"
        )
    }
  )

cancelled_provider_visits <- add_reporting_periods(
  provider_attributions_all_visits |>
    dplyr::filter(explicitly_cancelled_visit)
)

cancellations_provider <- cancelled_provider_visits |>
  dplyr::group_split(clinic_key, attributed_provider_key) |>
  purrr::map_dfr(
    function(provider_data) {
      provider_data |>
        summarize_cancellations_v3() |>
        dplyr::mutate(
          reporting_level = "Provider",
          reporting_level_order = 3L,
          cancellation_clinic = unique(provider_data$clinic_name)[1],
          cancellation_clinic_id = unique(provider_data$clinic_id)[1],
          cancellation_provider = unique(provider_data$attributed_provider_name)[1],
          cancellation_provider_id = unique(provider_data$attributed_provider_id)[1]
        )
    }
  )

cancellations_data <- dplyr::bind_rows(
  cancellations_portfolio,
  cancellations_clinic,
  cancellations_provider
) |>
  dplyr::arrange(
    reporting_period_order,
    reporting_period_date,
    reporting_level_order,
    cancellation_clinic,
    cancellation_provider,
    dplyr::desc(cancelled_visits)
  ) |>
  dplyr::transmute(
    `Reporting Period` = reporting_period,
    `Reporting Level` = reporting_level,
    `Clinic` = cancellation_clinic,
    `Clinic ID` = cancellation_clinic_id,
    `Provider/Resource` = cancellation_provider,
    `Provider/Resource ID` = cancellation_provider_id,
    `Cancellation Reason` = cancellation_reason_display,
    `Cancelled Visits` = cancelled_visits,
    `Cancelled Individual Visits` = cancelled_individual_visits,
    `Cancelled Group Visits` = cancelled_group_visits,
    `Affected Group Sessions` = affected_group_sessions,
    `Percentage of Cancellations` = percentage_of_cancellations
  )

cancellation_detail <- visits |>
  dplyr::filter(explicitly_cancelled_visit) |>
  dplyr::arrange(clinic_name, visit_date, appointment_time_seconds) |>
  dplyr::transmute(
    `Visit Record ID` = visit_record_id,
    `Source File` = source_file,
    `Source Row` = source_row,
    `Clinic` = clinic_name,
    `Clinic ID` = clinic_id,
    `Provider/Resource` = provider_resource,
    `Visit Type` = visit_type,
    `Type` = internal_type,
    `Visit Date` = visit_date,
    `Start Time` = appointment_time,
    `Cancellation Date` = cancellation_date,
    `Cancellation Reason` = cancellation_reason_display,
    `Appointment Status` = appt_status,
    `Group Classification` = group_visit_classification,
    `Group Session Outcome` = group_session_outcome,
    `Session ID` = session_id,
    `Data Quality Issue` = data_quality_issue
  )

completed_individual_visits <- visits |>
  dplyr::filter(
    group_visit_classification == "Individual visit",
    manager_status_category == "Completed",
    !is.na(start_hour)
  )

completed_group_sessions <- visits |>
  dplyr::filter(
    potential_group_visit,
    group_session_completed,
    !is.na(start_hour)
  ) |>
  dplyr::distinct(session_id, .keep_all = TRUE)

completed_individual_provider <- create_provider_attributions(
  completed_individual_visits
)

completed_group_session_provider <- create_provider_attributions(
  completed_group_sessions
)

build_chart_activity <- function(
    individual_data,
    group_data,
    chart_level_value,
    clinic_name_value,
    clinic_id_value,
    provider_name_value,
    provider_id_value
) {
  individual_summary <- add_reporting_periods(individual_data) |>
    dplyr::group_by(
      reporting_period,
      reporting_period_date,
      reporting_period_order,
      start_hour
    ) |>
    dplyr::summarise(
      completed_individual_visits = dplyr::n_distinct(visit_record_id),
      .groups = "drop"
    )

  group_summary <- add_reporting_periods(group_data) |>
    dplyr::group_by(
      reporting_period,
      reporting_period_date,
      reporting_period_order,
      start_hour
    ) |>
    dplyr::summarise(
      completed_group_sessions = dplyr::n_distinct(session_id),
      .groups = "drop"
    )

  dplyr::full_join(
    individual_summary,
    group_summary,
    by = c(
      "reporting_period",
      "reporting_period_date",
      "reporting_period_order",
      "start_hour"
    )
  ) |>
    dplyr::mutate(
      completed_individual_visits = dplyr::coalesce(
        completed_individual_visits,
        0L
      ),
      completed_group_sessions = dplyr::coalesce(
        completed_group_sessions,
        0L
      ),
      chart_level = chart_level_value,
      clinic_name = clinic_name_value,
      clinic_id = clinic_id_value,
      provider_name = provider_name_value,
      provider_id = provider_id_value
    )
}

chart_portfolio <- build_chart_activity(
  completed_individual_visits,
  completed_group_sessions,
  "Full Portfolio",
  portfolio_display_name,
  "FULL_PORTFOLIO",
  all_providers_display_name,
  all_providers_display_name
)

chart_clinic <- visits |>
  dplyr::distinct(clinic_key, clinic_name, clinic_id) |>
  purrr::pmap_dfr(
    function(clinic_key, clinic_name, clinic_id) {
      selected_clinic_key <- clinic_key
      build_chart_activity(
        completed_individual_visits |>
          dplyr::filter(.data$clinic_key == selected_clinic_key),
        completed_group_sessions |>
          dplyr::filter(.data$clinic_key == selected_clinic_key),
        "Full Clinic",
        clinic_name,
        clinic_id,
        all_providers_display_name,
        all_providers_display_name
      )
    }
  )

chart_provider_combinations <- provider_attributions_all_visits |>
  dplyr::distinct(
    clinic_key,
    clinic_name,
    clinic_id,
    attributed_provider_key,
    attributed_provider_name,
    attributed_provider_id
  )

chart_provider <- chart_provider_combinations |>
  purrr::pmap_dfr(
    function(
    clinic_key,
    clinic_name,
    clinic_id,
    attributed_provider_key,
    attributed_provider_name,
    attributed_provider_id
    ) {
      selected_clinic_key <- clinic_key
      selected_provider_key <- attributed_provider_key
      build_chart_activity(
        completed_individual_provider |>
          dplyr::filter(
            .data$clinic_key == selected_clinic_key,
            .data$attributed_provider_key == selected_provider_key
          ),
        completed_group_session_provider |>
          dplyr::filter(
            .data$clinic_key == selected_clinic_key,
            .data$attributed_provider_key == selected_provider_key
          ),
        "Provider",
        clinic_name,
        clinic_id,
        attributed_provider_name,
        attributed_provider_id
      )
    }
  )

chart_data_base <- dplyr::bind_rows(
  chart_portfolio,
  chart_clinic,
  chart_provider
) |>
  dplyr::mutate(
    chart_selection = dplyr::case_when(
      chart_level == "Full Portfolio" ~ portfolio_display_name,
      chart_level == "Full Clinic" ~ paste0(
        clinic_name,
        " | ",
        all_providers_display_name
      ),
      TRUE ~ paste0(clinic_name, " | ", provider_name)
    )
  )

chart_options <- chart_data_base |>
  dplyr::distinct(
    reporting_period,
    reporting_period_date,
    reporting_period_order,
    chart_level,
    clinic_name,
    clinic_id,
    provider_name,
    provider_id,
    chart_selection
  )

chart_data <- tidyr::crossing(
  chart_options,
  start_hour = 0:23
) |>
  dplyr::left_join(
    chart_data_base,
    by = c(
      "reporting_period",
      "reporting_period_date",
      "reporting_period_order",
      "chart_level",
      "clinic_name",
      "clinic_id",
      "provider_name",
      "provider_id",
      "chart_selection",
      "start_hour"
    )
  ) |>
  dplyr::mutate(
    completed_individual_visits = dplyr::coalesce(
      completed_individual_visits,
      0L
    ),
    completed_group_sessions = dplyr::coalesce(
      completed_group_sessions,
      0L
    ),
    total_service_activity =
      completed_individual_visits + completed_group_sessions,
    start_hour_label = format_hour_label(start_hour)
  ) |>
  dplyr::arrange(
    reporting_period_order,
    reporting_period_date,
    chart_selection,
    start_hour
  ) |>
  dplyr::transmute(
    `Reporting Period` = reporting_period,
    `Chart Level` = chart_level,
    `Clinic` = clinic_name,
    `Clinic ID` = clinic_id,
    `Provider/Resource` = provider_name,
    `Provider/Resource ID` = provider_id,
    `Chart Selection` = chart_selection,
    `Start Hour Number` = start_hour,
    `Start Hour` = start_hour_label,
    `Completed Individual Visits` = completed_individual_visits,
    `Completed Group Sessions` = completed_group_sessions,
    `Total Service Activity` = total_service_activity
  )

clinic_reference_output <- visits |>
  dplyr::distinct(
    clinic_name,
    clinic_id,
    clinic_key,
    appt_department
  ) |>
  dplyr::arrange(clinic_name, clinic_id) |>
  dplyr::transmute(
    `Clinic` = clinic_name,
    `Clinic ID` = clinic_id,
    `Clinic Key` = clinic_key,
    `Appt Department Source Value` = appt_department
  )

# ==============================================================================
# 27. CREATE DATA-QUALITY REPORTS
# ==============================================================================
#
# Creates:
#
#   - A visit-level Data Quality Detail worksheet
#   - A summary of important import and data-quality measures
#
# Records are retained whenever possible and flagged for review.
#
# ==============================================================================

data_quality_detail <- visits |>
  dplyr::filter(
    !is.na(data_quality_issue) |
      !group_rule_fields_complete
  ) |>
  dplyr::arrange(
    source_file,
    source_row
  ) |>
  dplyr::transmute(
    `Visit Record ID` =
      visit_record_id,

    `Source File` =
      source_file,

    `Source Row` =
      source_row,

    `Visit Date` =
      visit_date,

    `Start Time` =
      appointment_time,

    `Provider/Resource` =
      provider_resource,

    `External Visit Type` =
      external_visit_type_name,

    `Internal Visit Type` =
      internal_visit_type_name,

    `External Visit Type ID` =
      external_visit_type_id,

    `Internal Visit Type ID` =
      internal_visit_type_id,

    `Appointment Status` =
      appt_status,

    `Manager Status Category` =
      manager_status_category,

    `Data Quality Issue` =
      data_quality_issue
  )

potential_group_session_count <-
  dplyr::n_distinct(
    qualifying_visits$session_id[
      qualifying_visits$potential_group_visit
    ],
    na.rm = TRUE
  )

data_quality_summary <- tibble::tibble(
  Measure = c(
    "Input Excel files",
    "Total imported visit rows",
    "Visits before 4:00 PM",
    "Visits at 4:00 PM or later",
    "Visits across all valid time frames",
    "Visits with missing or invalid times",
    "Visits exactly at 4:00 PM",
    "Visits at 4:15 PM or later",
    "Inferred group visits at 4:00 PM or later",
    "Inferred group sessions at 4:00 PM or later",
    "Rows with missing or invalid visit dates",
    "Rows with missing providers/resources",
    "Rows with missing visit types",
    "Rows with missing MRNs",
    "Rows with mismatched Visit Type IDs",
    "Rows marked as possible duplicates",
    "Visits at 4:00 PM or later with Other status"
  ),

  Count = c(
    length(input_files),

    nrow(visits),

    nrow(visits_before_4_pm),

    nrow(visits_4_pm_or_later),

    nrow(visits_all_time_frames),

    sum(
      is.na(
        visits$appointment_time_seconds
      )
    ),

    sum(
      visits$time_category ==
        "4:00 PM to 4:14 PM",
      na.rm = TRUE
    ),

    sum(
      visits$time_category ==
        "4:15 PM or later",
      na.rm = TRUE
    ),

    sum(
      qualifying_visits$potential_group_visit,
      na.rm = TRUE
    ),

    potential_group_session_count,

    sum(
      is.na(
        visits$visit_date
      )
    ),

    sum(
      is.na(
        visits$provider_resource
      )
    ),

    sum(
      is.na(
        visits$visit_type_grouping_key
      )
    ),

    sum(
      is.na(
        visits$mrn
      )
    ),

    sum(
      visits$visit_type_id_mismatch,
      na.rm = TRUE
    ),

    sum(
      visits$possible_duplicate_appointment,
      na.rm = TRUE
    ),

    sum(
      qualifying_visits$manager_status_category ==
        "Other",
      na.rm = TRUE
    )
  )
)


# ==============================================================================
# 28. CREATE THE FILE-PROCESSING SUMMARY
# ==============================================================================
#
# Summarizes each imported workbook, including:
#
#   - Imported row count
#   - Earliest and latest visit dates
#   - Counts before and after 4:00 PM
#   - Data-quality issue counts
#
# ==============================================================================

file_summary <- visits |>
  dplyr::group_by(
    source_file
  ) |>
  dplyr::summarise(
    imported_rows =
      dplyr::n(),

    earliest_visit_date =
      if (
        all(
          is.na(visit_date)
        )
      ) {
        as.Date(NA)
      } else {
        min(
          visit_date,
          na.rm = TRUE
        )
      },

    latest_visit_date =
      if (
        all(
          is.na(visit_date)
        )
      ) {
        as.Date(NA)
      } else {
        max(
          visit_date,
          na.rm = TRUE
        )
      },

    visits_before_4_pm =
      sum(
        !is.na(appointment_time_seconds) &
          appointment_time_seconds <
          four_pm_seconds,
        na.rm = TRUE
      ),

    visits_4_pm_or_later =
      sum(
        overview_qualifying_visit,
        na.rm = TRUE
      ),

    visits_with_invalid_time =
      sum(
        is.na(
          appointment_time_seconds
        )
      ),

    inferred_group_visits_4_pm_or_later =
      sum(
        overview_qualifying_visit &
          potential_group_visit,
        na.rm = TRUE
      ),

    data_quality_issues =
      sum(
        !is.na(
          data_quality_issue
        ),
        na.rm = TRUE
      ),

    .groups =
      "drop"
  ) |>
  dplyr::arrange(
    source_file
  ) |>
  dplyr::transmute(
    `Source File` =
      source_file,

    `Imported Rows` =
      imported_rows,

    `Earliest Visit Date` =
      earliest_visit_date,

    `Latest Visit Date` =
      latest_visit_date,

    `Visits Before 4:00 PM` =
      visits_before_4_pm,

    `Visits at 4:00 PM or Later` =
      visits_4_pm_or_later,

    `Visits with Missing or Invalid Time` =
      visits_with_invalid_time,

    `Inferred Group Visits at 4:00 PM or Later` =
      inferred_group_visits_4_pm_or_later,

    `Data Quality Issues` =
      data_quality_issues
  )


# ==============================================================================
# 29. CREATE THE DATA DICTIONARY
# ==============================================================================
#
# Documents the business rules, status mappings, measures, privacy decisions,
# and interpretation cautions used by Version 3.
#
# ==============================================================================

data_dictionary <- tibble::tribble(
  ~Category, ~Term, ~Definition, ~Notes,

  "Reporting",
  "Reporting Period",
  "The month or combined period represented by an Overview or summary row.",
  "Monthly periods are derived from Visit Date rather than the source filename.",

  "Reporting",
  "All Months Combined",
  "Results calculated across every monthly input file included in the current analysis run.",
  "Additional months are included automatically when new workbooks are added to input_data.",

  "Time Frame",
  "Before 4:00 PM",
  "Appointments with a valid scheduled start time earlier than 4:00 PM.",
  "Appointments beginning exactly at 4:00 PM are excluded.",

  "Time Frame",
  "4:00 PM or Later",
  "Appointments with a valid scheduled start time at exactly 4:00 PM or any later time.",
  "This combines Exactly 4:00 PM and 4:15 PM or later.",

  "Time Frame",
  "All Time Frames",
  "All appointments with a valid scheduled start time.",
  "Calculated directly from the underlying visit records rather than by adding displayed summary rows.",

  "Time Frame",
  "4:00 PM to 4:14 PM",
  "An appointment whose scheduled start time is exactly 4:00 PM.",
  "Retained separately in the Monthly Comparison and detailed data.",

  "Time Frame",
  "4:15 PM or Later",
  "An appointment whose scheduled start time is 4:15 PM or any later time.",
  "This category excludes appointments beginning exactly at 4:00 PM.",

  "Time Frame",
  "Missing or Invalid Time",
  "An appointment whose start time is missing or cannot be interpreted.",
  "Excluded from the three Overview time frames but retained in Raw Data and Data Quality Detail.",

  "Clinic Counting",
  "Full Clinic",
  "The unique clinic-level count in which each original appointment is counted once.",
  "A joint visit involving multiple providers is counted only once in Full Clinic totals.",

  "Provider Attribution",
  "Provider/Resource",
  "The staff member, scheduling resource, or combination recorded in the source Provider/Resource field.",
  "The original source value is preserved in Raw Data and Provider Attribution.",

  "Provider Attribution",
  "Attributed Visit",
  "A provider-level participation count in which a visit is attributed once to each participating provider.",
  "Provider-attribution totals may exceed Full Clinic totals.",

  "Provider Attribution",
  "Joint Visit",
  "An appointment for which more than one provider or resource is identified in the source Provider/Resource field.",
  "Counted once for Full Clinic and once for each individually identified provider.",

  "Provider Attribution",
  "Provider/Resource ID",
  "The identifier shown in square brackets after a provider or resource name.",
  "Used as the preferred provider grouping key.",

  "Provider Attribution",
  "Provider Parse Method",
  "A description of how an individual provider was extracted from the original Provider/Resource field.",
  "Provider parsing can be reviewed on the Provider Attribution worksheet.",

  "Provider Attribution",
  "Provider Totals Versus Full Clinic",
  "Provider rows represent attributed participation, while Full Clinic rows represent unique appointments.",
  "Do not add provider rows together to calculate clinic totals.",

  "Visit Classification",
  "Inferred Group Visit",
  "A patient appointment belonging to a cluster containing at least two distinct MRNs with the same date, time, original Provider/Resource value, and visit type.",
  "This classification is inferred from scheduling patterns and is not an explicit source-system group indicator.",

  "Visit Classification",
  "Individual Visit",
  "An appointment that does not meet the inferred group rule and has all fields required to apply the rule.",
  "This is a patient-level appointment measure.",

  "Visit Classification",
  "Unclassified/Review Visit",
  "An appointment for which the group rule cannot be fully applied because a required field is missing or invalid.",
  "Retained and displayed separately rather than automatically classified as individual.",

  "Visit Classification",
  "Inferred Group Session",
  "One distinct scheduling event defined by date, start time, original Provider/Resource value, and visit type, containing at least two distinct MRNs.",
  "One group session may contain multiple patient-level group visits.",

  "Visit Classification",
  "Distinct Patients in Session",
  "The number of unique MRNs associated with the same inferred group session.",
  "Distinct patients are counted using MRN rather than source-row count.",

  "Status",
  "Appointment Status",
  "The original appointment status recorded in the source workbook.",
  "Original values remain visible in Status Detail and Raw Data.",

  "Status",
  "Completed",
  "The manager category assigned to source statuses such as Comp, Complete, or Completed.",
  "The known CPC status Comp maps to Completed.",

  "Status",
  "Cancelled/No-Show",
  "The manager category combining cancelled appointments and no-shows.",
  "Known CPC statuses include Can and No Show.",

  "Status",
  "Other Status",
  "A source status that is not mapped to Completed or Cancelled/No-Show.",
  "For example, Left remains in Other Status.",

  "Individual Visit Outcome",
  "Completed Individual Visit",
  "An appointment classified as an individual visit whose manager status is Completed.",
  "A primary manager-facing Overview measure.",

  "Individual Visit Outcome",
  "Cancelled/No-Show Individual Visit",
  "An appointment classified as an individual visit whose manager status is Cancelled/No-Show.",
  "Cancellations and no-shows are combined in the Overview.",

  "Individual Visit Outcome",
  "Other Individual Visit",
  "An individual visit whose manager status is Other.",
  "Allows the individual visit outcomes to reconcile to Individual Visits.",

  "Group Visit Outcome",
  "Completed Group Visit",
  "A patient-level appointment within an inferred group whose manager status is Completed.",
  "One completed session may contain several completed group visits.",

  "Group Visit Outcome",
  "Cancelled/No-Show Group Visit",
  "A patient-level appointment within an inferred group whose manager status is Cancelled/No-Show.",
  "A delivered group may still contain cancelled or no-show patient visits.",

  "Group Session Outcome",
  "Completed Group Session",
  "An inferred group session containing at least one completed patient-level visit.",
  "A session remains completed even if some scheduled patients cancelled or did not attend.",

  "Group Session Outcome",
  "Cancelled/No-Show Group Session",
  "An inferred group session with no completed visits where every patient-level visit is Cancelled/No-Show.",
  "A session is not treated as cancelled because only one participant cancelled.",

  "Group Session Outcome",
  "Other/Review Group Session",
  "An inferred group session with no completed visits and at least one patient-level visit with an Other status.",
  "These sessions may require review.",

  "Overview Measure",
  "Total Visits",
  "The number of unique appointment records included in the Overview row.",
  "Provider rows show attributed visits, while Full Clinic rows show unique clinic appointments.",

  "Overview Measure",
  "Individual Visits",
  "The number of appointment records classified as individual visits.",
  "Completed, Cancelled/No-Show, and Other Individual Visits should reconcile to this measure.",

  "Overview Measure",
  "Inferred Group Visits",
  "The number of patient-level appointment records classified as inferred group visits.",
  "This is not the number of group sessions.",

  "Overview Measure",
  "Inferred Group Sessions",
  "The number of distinct inferred group scheduling events.",
  "One session may contain multiple patient visits.",

  "Overview Measure",
  "Unclassified/Review Visits",
  "The number of appointments for which group-versus-individual classification could not be completed.",
  "Shown separately so Total Visits remains reconcilable.",

  "Rate",
  "Individual Completion Rate",
  "Completed Individual Visits divided by Individual Visits.",
  "Blank when there are no individual visits in the row.",

  "Rate",
  "Individual Cancelled/No-Show Rate",
  "Cancelled/No-Show Individual Visits divided by Individual Visits.",
  "Blank when there are no individual visits in the row.",

  "Rate",
  "Group Session Completion Rate",
  "Completed Group Sessions divided by Inferred Group Sessions.",
  "Blank when there are no inferred group sessions in the row.",

  "Rate",
  "Group Session Cancelled/No-Show Rate",
  "Cancelled/No-Show Group Sessions divided by Inferred Group Sessions.",
  "Blank when there are no inferred group sessions in the row.",

  "Status Detail",
  "Inferred Group Sessions by Appointment Status",
  "The number of distinct group sessions containing at least one patient visit with the displayed original appointment status.",
  "Do not sum this column across statuses because one session can contain multiple statuses.",

  "Source Field",
  "MRN",
  "The patient medical record number.",
  "Used internally to identify distinct patients but excluded from every output worksheet.",

  "Source Field",
  "Patient",
  "The patient name recorded in the source workbook.",
  "Excluded from every output worksheet.",

  "Source Field",
  "Appt Department",
  "The appointment department imported from the source Department column.",
  "Used to derive Clinic, Clinic ID, clinic grouping, and portfolio reporting.",

  "Source Field",
  "Provider Department",
  "The staff organizational department imported from the source Dept column.",
  "Retained in Raw Data for reference but not currently used in calculations.",

  "Source Field",
  "Provider/Resource Source Field",
  "The original provider or resource value from the source workbook.",
  "May contain one provider, multiple providers, or another scheduling resource.",

  "Source Field",
  "Visit Type",
  "The external or primary visit-type description and its bracketed identifier.",
  "Used in group inference and quality checks.",

  "Source Field",
  "Type",
  "The internal-facing visit-type description and its bracketed identifier.",
  "The displayed name may differ from Visit Type.",

  "Source Field",
  "Visit Date",
  "The scheduled date of the appointment.",
  "Determines the reporting month.",

  "Source Field",
  "Time",
  "The scheduled start time of the appointment.",
  "Standardized before time-period classification.",

  "Source Field",
  "Appt Status",
  "The original appointment status from the source workbook.",
  "Mapped into a manager status category while remaining available for review.",

  "Source Field",
  "Canc Date",
  "The cancellation date recorded in the source workbook.",
  "May be blank when the appointment was not cancelled.",

  "Source Field",
  "Canc Reason",
  "The cancellation reason recorded in the source workbook.",
  "May be blank when no cancellation reason applies.",

  "Audit Field",
  "Source File",
  "The monthly workbook from which the appointment was imported.",
  "Supports reconciliation across input files.",

  "Audit Field",
  "Source Row",
  "The approximate source-workbook row number, including the header row.",
  "Can help locate a record in the original workbook.",

  "Audit Field",
  "Visit Record ID",
  "A sequential identifier assigned to each imported appointment during the analysis run.",
  "May change if source files or row order change.",

  "Derived Field",
  "Visit Month",
  "The calendar month derived from Visit Date.",
  "Displayed as Month Year.",

  "Derived Field",
  "Time Category",
  "The detailed scheduled-time category assigned by the tool.",
  "Values include Before 4:00 PM, Exactly 4:00 PM, 4:15 PM or later, and Missing or invalid appointment time.",

  "Derived Field",
  "Manager Status Category",
  "The reporting category derived from the original appointment status.",
  "Values are Completed, Cancelled/No-Show, and Other.",

  "Derived Field",
  "Session ID",
  "An analytical identifier representing date, time, original provider combination, and visit type.",
  "Created by the tool and not sourced from the clinical system.",

  "Derived Field",
  "Group Session Outcome",
  "The overall outcome assigned to an inferred group session.",
  "Values are Completed, Cancelled/No-Show, or Other/Review Group Session.",

  "Data Quality",
  "Possible Duplicate Appointment",
  "The same MRN appears more than once for the same date, time, provider combination, and visit type.",
  "Possible duplicates are retained and flagged rather than automatically removed.",

  "Data Quality",
  "Visit Type ID Mismatch",
  "The bracketed identifier in Visit Type differs from the identifier in Type.",
  "Does not automatically remove a visit from the analysis.",

  "Data Quality",
  "Data Quality Issue",
  "The first identified major quality issue for a visit record.",
  "Records are retained whenever possible and flagged for review.",

  "Privacy",
  "Operational Raw Data",
  "The Raw Data worksheet contains one row per appointment while excluding MRN and Patient.",
  "The worksheet remains operationally sensitive and should be handled according to organizational privacy requirements.",

  "Worksheet Scope",
  "Overview",
  "Manager-facing clinic and provider summary for Before 4:00 PM, 4:00 PM or Later, and All Time Frames.",
  "Includes All Months Combined and individual reporting months.",

  "Worksheet Scope",
  "Monthly Comparison",
  "Clinic-level monthly analysis of appointments at 4:00 PM or later.",
  "Retains the original late-visit project scope.",

  "Worksheet Scope",
  "Provider Detail",
  "Provider-attribution detail for appointments at 4:00 PM or later.",
  "Joint visits may appear once for each participating provider.",

  "Worksheet Scope",
  "Status Detail",
  "Original appointment-status analysis for visits at 4:00 PM or later.",
  "Includes All Months Combined and monthly results.",

  "Worksheet Scope",
  "Group Sessions",
  "Session-level listing of inferred groups at 4:00 PM or later.",
  "Each row represents one inferred group session.",

  "Worksheet Scope",
  "Provider Attribution",
  "Audit worksheet showing how original Provider/Resource values were assigned to individual providers.",
  "Used to validate joint-provider parsing.",

  "Worksheet Scope",
  "Raw Data",
  "Appointment-level dataset with direct patient identifiers omitted covering all imported appointments.",
  "Contains one row per original appointment.",

  "Worksheet Scope",
  "Data Quality",
  "Summary and detail worksheets containing records that may require review.",
  "Quality flags do not automatically remove visits."
  ,
  "Reporting Level",
  "Full Portfolio",
  "A unique appointment-level summary across all clinics in the source data.",
  "Each appointment is counted once at the portfolio level.",

  "Clinic",
  "Clinic",
  "The clinic or program derived from the source Appt Department value.",
  "The bracketed Appt Department ID is used as the preferred clinic identifier.",

  "Time Frame",
  "8:00 AM to 4:00 PM",
  "Appointments starting at or after 8:00 AM and before 4:00 PM.",
  "Version 3 time frames overlap and should not be added together.",

  "Time Frame",
  "9:00 AM to 5:00 PM",
  "Appointments starting at or after 9:00 AM and before 5:00 PM.",
  "Version 3 time frames overlap and should not be added together.",

  "Time Frame",
  "8:00 AM to 9:00 AM",
  "Appointments starting at or after 8:00 AM and before 9:00 AM.",
  "A visit at exactly 9:00 AM is excluded.",

  "Time Frame",
  "4:30 PM or Later",
  "Appointments starting at exactly 4:30 PM or later.",
  "This is a subset of 4:00 PM or Later.",

  "Patient Measure",
  "Unique Patients Scheduled",
  "Distinct non-missing MRNs with any appointment represented in the report row.",
  "MRNs are used internally and are not exported.",

  "Patient Measure",
  "Unique Patients Served",
  "Distinct non-missing MRNs with at least one completed patient-level visit.",
  "MRNs are used internally and are not exported.",

  "Cancellation",
  "Explicit Cancellation",
  "A visit whose source status specifically indicates cancellation.",
  "No-shows remain in combined status measures but are excluded from cancellation-reason analysis.",

  "Service Activity",
  "Total Service Activity",
  "Completed individual visits plus completed group sessions by scheduled start hour.",
  "Each group session counts once regardless of participant count."

)


# ==============================================================================
# 30. DEFINE GENERAL EXCEL FORMATTING
# ==============================================================================
#
# Creates a reusable function that applies consistent:
#
#   - Fonts
#   - Header formatting
#   - Filters
#   - Frozen headers
#   - Automatic column widths
#
# ==============================================================================

format_excel_sheet <- function(
    workbook,
    sheet_name,
    data
) {

  if (ncol(data) == 0) {
    return(invisible(NULL))
  }

  header_style <- openxlsx::createStyle(
    fontName = "Aptos",
    fontSize = 11,
    fontColour = "#FFFFFF",
    fgFill = "#1F4E78",
    textDecoration = "bold",
    halign = "center",
    valign = "center",
    wrapText = TRUE,
    border = "Bottom",
    borderColour = "#FFFFFF"
  )

  body_style <- openxlsx::createStyle(
    fontName = "Aptos",
    fontSize = 10,
    valign = "top"
  )

  openxlsx::addStyle(
    wb = workbook,
    sheet = sheet_name,
    style = header_style,
    rows = 1,
    cols = seq_len(
      ncol(data)
    ),
    gridExpand = TRUE,
    stack = TRUE
  )

  if (nrow(data) > 0) {

    openxlsx::addStyle(
      wb = workbook,
      sheet = sheet_name,
      style = body_style,
      rows = 2:(nrow(data) + 1),
      cols = seq_len(
        ncol(data)
      ),
      gridExpand = TRUE,
      stack = TRUE
    )

    openxlsx::addFilter(
      wb = workbook,
      sheet = sheet_name,
      rows = 1,
      cols = seq_len(
        ncol(data)
      )
    )
  }

  openxlsx::freezePane(
    wb = workbook,
    sheet = sheet_name,
    firstRow = TRUE
  )

  openxlsx::setColWidths(
    wb = workbook,
    sheet = sheet_name,
    cols = seq_len(
      ncol(data)
    ),
    widths = "auto"
  )

  long_text_columns <- which(
    names(data) %in% c(
      "Canc Reason",
      "Data Quality Issue",
      "Appointment Statuses",
      "Source Files",
      "Original Provider/Resource",
      "Provider Parse Method",
      "Definition",
      "Notes"
    )
  )

  if (
    length(long_text_columns) > 0 &&
    nrow(data) > 0
  ) {

    wrap_style <- openxlsx::createStyle(
      wrapText = TRUE,
      valign = "top"
    )

    openxlsx::setColWidths(
      wb = workbook,
      sheet = sheet_name,
      cols = long_text_columns,
      widths = 35
    )

    openxlsx::addStyle(
      wb = workbook,
      sheet = sheet_name,
      style = wrap_style,
      rows = 2:(nrow(data) + 1),
      cols = long_text_columns,
      gridExpand = TRUE,
      stack = TRUE
    )
  }

  invisible(NULL)
}


# ==============================================================================
# 31. CREATE AND POPULATE THE OUTPUT WORKBOOK
# ==============================================================================
#
# Creates the output workbook and adds all manager, detail, audit, raw-data,
# quality-control, and documentation worksheets.
#
# ==============================================================================

workbook <- openxlsx::createWorkbook(
  creator = "Clinic Capacity Planner 0.1 - original V3 analysis"
)

worksheet_data <- list(
  "Overview" =
    overview_data,

  "Cancellations" =
    cancellations_data,

  "Cancellation Detail" =
    cancellation_detail,

  "Unique Patients" =
    unique_patients_data,

  "Clinic Reference" =
    clinic_reference_output,

  "Chart Data" =
    chart_data,

  "Monthly Comparison" =
    monthly_comparison,

  "Provider Detail" =
    provider_detail,

  "Status Detail" =
    status_detail,

  "Group Sessions" =
    group_sessions,

  "Provider Attribution" =
    provider_attribution_audit,

  "Raw Data" =
    raw_data,

  "Data Quality Summary" =
    data_quality_summary,

  "Data Quality Detail" =
    data_quality_detail,

  "Files Processed" =
    file_summary,

  "Data Dictionary" =
    data_dictionary
)

decision_result <- cc_analyse(visits, decision_config, full_capacity, full_scenarios, full_hours_plans)
worksheet_data <- c(cc_tables(decision_result), worksheet_data)

purrr::walk2(
  names(
    worksheet_data
  ),
  worksheet_data,
  function(
    sheet_name,
    sheet_data
  ) {

    openxlsx::addWorksheet(
      wb = workbook,
      sheetName = sheet_name,
      gridLines = FALSE
    )

    if (nrow(sheet_data) == 0) {

      displayed_data <- tibble::tibble(
        Message =
          "No records met the criteria for this worksheet."
      )

    } else {

      displayed_data <- sheet_data
    }

    openxlsx::writeData(
      wb = workbook,
      sheet = sheet_name,
      x = displayed_data,
      withFilter = FALSE,
      keepNA = FALSE
    )

    format_excel_sheet(
      workbook = workbook,
      sheet_name = sheet_name,
      data = displayed_data
    )
  }
)



# ==============================================================================
# 31A. CREATE THE INTERACTIVE SERVICE ACTIVITY WORKSHEET
# ==============================================================================
#
# Creates dropdown controls for reporting period and portfolio, clinic, or
# provider selection. The hourly display uses completed individual visits and
# completed group sessions. Excel data bars provide a graph-like hourly view.
#
# ==============================================================================

chart_reporting_period_options <- chart_data |>
  dplyr::distinct(`Reporting Period`) |>
  dplyr::mutate(
    sort_date = dplyr::case_when(
      `Reporting Period` == "All Months Combined" ~ as.Date("1900-01-01"),
      TRUE ~ suppressWarnings(
        as.Date(paste0("01 ", `Reporting Period`), format = "%d %B %Y")
      )
    )
  ) |>
  dplyr::arrange(sort_date) |>
  dplyr::pull(`Reporting Period`)

chart_selection_options <- chart_data |>
  dplyr::distinct(`Chart Selection`) |>
  dplyr::arrange(`Chart Selection`) |>
  dplyr::pull(`Chart Selection`)

openxlsx::addWorksheet(
  wb = workbook,
  sheetName = "Chart Lists",
  gridLines = FALSE
)

chart_list_length <- max(
  length(chart_reporting_period_options),
  length(chart_selection_options)
)

chart_lists <- tibble::tibble(
  `Reporting Period Options` = c(
    chart_reporting_period_options,
    rep(NA_character_, chart_list_length - length(chart_reporting_period_options))
  ),
  `Chart Selection Options` = c(
    chart_selection_options,
    rep(NA_character_, chart_list_length - length(chart_selection_options))
  )
)

openxlsx::writeData(
  wb = workbook,
  sheet = "Chart Lists",
  x = chart_lists
)

openxlsx::addWorksheet(
  wb = workbook,
  sheetName = "Service Activity Chart",
  gridLines = FALSE
)

openxlsx::mergeCells(
  wb = workbook,
  sheet = "Service Activity Chart",
  cols = 1:8,
  rows = 1
)

openxlsx::writeData(
  wb = workbook,
  sheet = "Service Activity Chart",
  x = "Service Activity Throughout the Day",
  startCol = 1,
  startRow = 1,
  colNames = FALSE
)

chart_title_style <- openxlsx::createStyle(
  fontName = "Aptos Display",
  fontSize = 18,
  fontColour = "#FFFFFF",
  fgFill = "#1F4E78",
  textDecoration = "bold"
)

openxlsx::addStyle(
  wb = workbook,
  sheet = "Service Activity Chart",
  style = chart_title_style,
  rows = 1,
  cols = 1:8,
  gridExpand = TRUE,
  stack = TRUE
)

openxlsx::writeData(
  wb = workbook,
  sheet = "Service Activity Chart",
  x = "Reporting Period",
  startCol = 1,
  startRow = 4,
  colNames = FALSE
)

openxlsx::writeData(
  wb = workbook,
  sheet = "Service Activity Chart",
  x = "Chart Selection",
  startCol = 1,
  startRow = 5,
  colNames = FALSE
)

openxlsx::writeData(
  wb = workbook,
  sheet = "Service Activity Chart",
  x = "All Months Combined",
  startCol = 2,
  startRow = 4,
  colNames = FALSE
)

openxlsx::writeData(
  wb = workbook,
  sheet = "Service Activity Chart",
  x = portfolio_display_name,
  startCol = 2,
  startRow = 5,
  colNames = FALSE
)

if (length(chart_reporting_period_options) > 0) {
  openxlsx::dataValidation(
    wb = workbook,
    sheet = "Service Activity Chart",
    cols = 2,
    rows = 4,
    type = "list",
    value = paste0(
      "'Chart Lists'!$A$2:$A$",
      length(chart_reporting_period_options) + 1
    )
  )
}

if (length(chart_selection_options) > 0) {
  openxlsx::dataValidation(
    wb = workbook,
    sheet = "Service Activity Chart",
    cols = 2,
    rows = 5,
    type = "list",
    value = paste0(
      "'Chart Lists'!$B$2:$B$",
      length(chart_selection_options) + 1
    )
  )
}

chart_headers <- c(
  "Start Hour",
  "Completed Individual Visits",
  "Completed Group Sessions",
  "Total Service Activity"
)

openxlsx::writeData(
  wb = workbook,
  sheet = "Service Activity Chart",
  x = t(chart_headers),
  startCol = 1,
  startRow = 8,
  colNames = FALSE
)

openxlsx::writeData(
  wb = workbook,
  sheet = "Service Activity Chart",
  x = tibble::tibble(
    `Start Hour` = format_hour_label(0:23),
    `Completed Individual Visits` = rep(0L, 24),
    `Completed Group Sessions` = rep(0L, 24),
    `Total Service Activity` = rep(0L, 24)
  ),
  startCol = 1,
  startRow = 9,
  colNames = FALSE
)

chart_data_last_row <- nrow(chart_data) + 1

for (chart_row in 9:32) {
  hour_number <- chart_row - 9

  individual_formula <- paste0(
    "SUMIFS('Chart Data'!$J$2:$J$", chart_data_last_row,
    ",'Chart Data'!$A$2:$A$", chart_data_last_row, ",$B$4,",
    "'Chart Data'!$G$2:$G$", chart_data_last_row, ",$B$5,",
    "'Chart Data'!$H$2:$H$", chart_data_last_row, ",", hour_number, ")"
  )

  group_formula <- paste0(
    "SUMIFS('Chart Data'!$K$2:$K$", chart_data_last_row,
    ",'Chart Data'!$A$2:$A$", chart_data_last_row, ",$B$4,",
    "'Chart Data'!$G$2:$G$", chart_data_last_row, ",$B$5,",
    "'Chart Data'!$H$2:$H$", chart_data_last_row, ",", hour_number, ")"
  )

  openxlsx::writeFormula(
    wb = workbook,
    sheet = "Service Activity Chart",
    x = individual_formula,
    startCol = 2,
    startRow = chart_row
  )

  openxlsx::writeFormula(
    wb = workbook,
    sheet = "Service Activity Chart",
    x = group_formula,
    startCol = 3,
    startRow = chart_row
  )

  openxlsx::writeFormula(
    wb = workbook,
    sheet = "Service Activity Chart",
    x = paste0("B", chart_row, "+C", chart_row),
    startCol = 4,
    startRow = chart_row
  )
}

openxlsx::conditionalFormatting(
  wb = workbook,
  sheet = "Service Activity Chart",
  cols = 2,
  rows = 9:32,
  type = "dataBar",
  style = "#5B9BD5"
)

openxlsx::conditionalFormatting(
  wb = workbook,
  sheet = "Service Activity Chart",
  cols = 3,
  rows = 9:32,
  type = "dataBar",
  style = "#70AD47"
)

openxlsx::conditionalFormatting(
  wb = workbook,
  sheet = "Service Activity Chart",
  cols = 4,
  rows = 9:32,
  type = "dataBar",
  style = "#ED7D31"
)

openxlsx::setColWidths(
  wb = workbook,
  sheet = "Service Activity Chart",
  cols = 1,
  widths = 15
)

openxlsx::setColWidths(
  wb = workbook,
  sheet = "Service Activity Chart",
  cols = 2:4,
  widths = 27
)

# ==============================================================================
# 32. FORMAT THE OVERVIEW WORKSHEET
# ==============================================================================
#
# Makes the manager-facing Overview easier to read by:
#
#   - Highlighting Full Clinic rows
#   - Colour-coding the three time frames
#   - Formatting counts and rates
#   - Applying conditional formatting to performance rates
#   - Freezing identifying columns
#
# ==============================================================================

if (nrow(overview_data) > 0) {

  overview_header_style <-
    openxlsx::createStyle(
      fontName = "Aptos",
      fontSize = 11,
      fontColour = "#FFFFFF",
      fgFill = "#1F4E78",
      textDecoration = "bold",
      halign = "center",
      valign = "center",
      wrapText = TRUE,
      border = "Bottom",
      borderColour = "#FFFFFF"
    )

  full_clinic_style <-
    openxlsx::createStyle(
      fontName = "Aptos",
      fontSize = 10,
      fontColour = "#FFFFFF",
      fgFill = "#4472C4",
      textDecoration = "bold",
      valign = "center"
    )

  before_4_pm_style <-
    openxlsx::createStyle(
      fontName = "Aptos",
      fontSize = 10,
      fgFill = "#F2F2F2",
      valign = "center"
    )

  four_pm_or_later_style <-
    openxlsx::createStyle(
      fontName = "Aptos",
      fontSize = 10,
      fgFill = "#D9EAF7",
      valign = "center"
    )

  all_time_frames_style <-
    openxlsx::createStyle(
      fontName = "Aptos",
      fontSize = 10,
      fgFill = "#E2F0D9",
      textDecoration = "bold",
      valign = "center"
    )

  count_style <-
    openxlsx::createStyle(
      numFmt = "0",
      halign = "center",
      valign = "center"
    )

  percentage_style <-
    openxlsx::createStyle(
      numFmt = "0.0%",
      halign = "center",
      valign = "center"
    )

  wrap_style <-
    openxlsx::createStyle(
      wrapText = TRUE,
      valign = "center"
    )

  openxlsx::addStyle(
    wb = workbook,
    sheet = "Overview",
    style = overview_header_style,
    rows = 1,
    cols = seq_len(
      ncol(overview_data)
    ),
    gridExpand = TRUE,
    stack = TRUE
  )

  full_clinic_rows <- which(
    overview_data$`Provider/Resource` ==
      "Full Clinic"
  ) + 1

  before_4_pm_provider_rows <- which(
    overview_data$`Time Frame` ==
      "Before 4:00 PM" &
      overview_data$`Reporting Level` ==
      "Provider"
  ) + 1

  four_pm_or_later_provider_rows <- which(
    overview_data$`Time Frame` ==
      "4:00 PM or Later" &
      overview_data$`Reporting Level` ==
      "Provider"
  ) + 1

  all_time_frame_provider_rows <- which(
    overview_data$`Time Frame` ==
      "Full Day" &
      overview_data$`Reporting Level` ==
      "Provider"
  ) + 1

  if (length(full_clinic_rows) > 0) {

    openxlsx::addStyle(
      wb = workbook,
      sheet = "Overview",
      style = full_clinic_style,
      rows = full_clinic_rows,
      cols = seq_len(
        ncol(overview_data)
      ),
      gridExpand = TRUE,
      stack = TRUE
    )
  }

  if (
    length(
      before_4_pm_provider_rows
    ) > 0
  ) {

    openxlsx::addStyle(
      wb = workbook,
      sheet = "Overview",
      style = before_4_pm_style,
      rows = before_4_pm_provider_rows,
      cols = seq_len(
        ncol(overview_data)
      ),
      gridExpand = TRUE,
      stack = TRUE
    )
  }

  if (
    length(
      four_pm_or_later_provider_rows
    ) > 0
  ) {

    openxlsx::addStyle(
      wb = workbook,
      sheet = "Overview",
      style = four_pm_or_later_style,
      rows = four_pm_or_later_provider_rows,
      cols = seq_len(
        ncol(overview_data)
      ),
      gridExpand = TRUE,
      stack = TRUE
    )
  }

  if (
    length(
      all_time_frame_provider_rows
    ) > 0
  ) {

    openxlsx::addStyle(
      wb = workbook,
      sheet = "Overview",
      style = all_time_frames_style,
      rows = all_time_frame_provider_rows,
      cols = seq_len(
        ncol(overview_data)
      ),
      gridExpand = TRUE,
      stack = TRUE
    )
  }

  count_columns <- which(
    names(overview_data) %in% c(
      "Total Visits",
      "Individual Visits",
      "Completed Individual Visits",
      "Cancelled/No-Show Individual Visits",
      "Other Individual Visits",
      "Inferred Group Visits",
      "Completed Group Visits",
      "Cancelled/No-Show Group Visits",
      "Inferred Group Sessions",
      "Completed Group Sessions",
      "Cancelled/No-Show Group Sessions",
      "Other/Review Group Sessions",
      "Unclassified/Review Visits"
    )
  )

  percentage_columns <- which(
    names(overview_data) %in% c(
      "Individual Completion Rate",
      "Individual Cancelled/No-Show Rate",
      "Group Session Completion Rate",
      "Group Session Cancelled/No-Show Rate"
    )
  )

  if (length(count_columns) > 0) {

    openxlsx::addStyle(
      wb = workbook,
      sheet = "Overview",
      style = count_style,
      rows = 2:(nrow(overview_data) + 1),
      cols = count_columns,
      gridExpand = TRUE,
      stack = TRUE
    )
  }

  if (length(percentage_columns) > 0) {

    openxlsx::addStyle(
      wb = workbook,
      sheet = "Overview",
      style = percentage_style,
      rows = 2:(nrow(overview_data) + 1),
      cols = percentage_columns,
      gridExpand = TRUE,
      stack = TRUE
    )
  }

  openxlsx::addStyle(
    wb = workbook,
    sheet = "Overview",
    style = wrap_style,
    rows = 1:(nrow(overview_data) + 1),
    cols = seq_len(
      ncol(overview_data)
    ),
    gridExpand = TRUE,
    stack = TRUE
  )

  reporting_period_column <- which(
    names(overview_data) ==
      "Reporting Period"
  )

  provider_column <- which(
    names(overview_data) ==
      "Provider/Resource"
  )

  provider_id_column <- which(
    names(overview_data) ==
      "Provider/Resource ID"
  )

  time_frame_column <- which(
    names(overview_data) ==
      "Time Frame"
  )

  if (
    length(
      reporting_period_column
    ) == 1
  ) {

    openxlsx::setColWidths(
      wb = workbook,
      sheet = "Overview",
      cols = reporting_period_column,
      widths = 21
    )
  }

  if (length(provider_column) == 1) {

    openxlsx::setColWidths(
      wb = workbook,
      sheet = "Overview",
      cols = provider_column,
      widths = 29
    )
  }

  if (
    length(
      provider_id_column
    ) == 1
  ) {

    openxlsx::setColWidths(
      wb = workbook,
      sheet = "Overview",
      cols = provider_id_column,
      widths = 20
    )
  }

  if (
    length(
      time_frame_column
    ) == 1
  ) {

    openxlsx::setColWidths(
      wb = workbook,
      sheet = "Overview",
      cols = time_frame_column,
      widths = 20
    )
  }

  measure_columns <- setdiff(
    seq_len(
      ncol(overview_data)
    ),
    c(
      reporting_period_column,
      provider_column,
      provider_id_column,
      time_frame_column
    )
  )

  if (length(measure_columns) > 0) {

    openxlsx::setColWidths(
      wb = workbook,
      sheet = "Overview",
      cols = measure_columns,
      widths = 19
    )
  }

  openxlsx::setRowHeights(
    wb = workbook,
    sheet = "Overview",
    rows = 1,
    heights = 48
  )

  if (nrow(overview_data) > 0) {

    openxlsx::setRowHeights(
      wb = workbook,
      sheet = "Overview",
      rows = 2:(nrow(overview_data) + 1),
      heights = 30
    )
  }

  rate_formatting <- list(
    list(
      column =
        "Individual Completion Rate",

      colours =
        c(
          "#F8696B",
          "#FFEB84",
          "#63BE7B"
        )
    ),

    list(
      column =
        "Individual Cancelled/No-Show Rate",

      colours =
        c(
          "#63BE7B",
          "#FFEB84",
          "#F8696B"
        )
    ),

    list(
      column =
        "Group Session Completion Rate",

      colours =
        c(
          "#F8696B",
          "#FFEB84",
          "#63BE7B"
        )
    ),

    list(
      column =
        "Group Session Cancelled/No-Show Rate",

      colours =
        c(
          "#63BE7B",
          "#FFEB84",
          "#F8696B"
        )
    )
  )

  purrr::walk(
    rate_formatting,
    function(formatting_rule) {

      target_column <- which(
        names(overview_data) ==
          formatting_rule$column
      )

      if (length(target_column) == 1) {

        openxlsx::conditionalFormatting(
          wb = workbook,
          sheet = "Overview",
          cols = target_column,
          rows = 2:(nrow(overview_data) + 1),
          type = "colourScale",
          style = formatting_rule$colours
        )
      }
    }
  )

  openxlsx::freezePane(
    wb = workbook,
    sheet = "Overview",
    firstActiveRow = 2,
    firstActiveCol = 5
  )
}


# ==============================================================================
# 33. FORMAT THE DATA DICTIONARY
# ==============================================================================
#
# Applies wrapped text, category highlighting, and wide definition columns to
# make the methodology and business rules easy to review.
#
# ==============================================================================

if (nrow(data_dictionary) > 0) {

  dictionary_body_style <-
    openxlsx::createStyle(
      fontName = "Aptos",
      fontSize = 10,
      valign = "top",
      wrapText = TRUE
    )

  dictionary_category_style <-
    openxlsx::createStyle(
      fontName = "Aptos",
      fontSize = 10,
      textDecoration = "bold",
      fgFill = "#D9EAF7",
      valign = "top",
      wrapText = TRUE
    )

  openxlsx::addStyle(
    wb = workbook,
    sheet = "Data Dictionary",
    style = dictionary_body_style,
    rows = 2:(nrow(data_dictionary) + 1),
    cols = 1:ncol(data_dictionary),
    gridExpand = TRUE,
    stack = TRUE
  )

  openxlsx::addStyle(
    wb = workbook,
    sheet = "Data Dictionary",
    style = dictionary_category_style,
    rows = 2:(nrow(data_dictionary) + 1),
    cols = 1,
    gridExpand = TRUE,
    stack = TRUE
  )

  openxlsx::setColWidths(
    wb = workbook,
    sheet = "Data Dictionary",
    cols = 1,
    widths = 22
  )

  openxlsx::setColWidths(
    wb = workbook,
    sheet = "Data Dictionary",
    cols = 2,
    widths = 34
  )

  openxlsx::setColWidths(
    wb = workbook,
    sheet = "Data Dictionary",
    cols = 3,
    widths = 70
  )

  openxlsx::setColWidths(
    wb = workbook,
    sheet = "Data Dictionary",
    cols = 4,
    widths = 65
  )

  openxlsx::setRowHeights(
    wb = workbook,
    sheet = "Data Dictionary",
    rows = 2:(nrow(data_dictionary) + 1),
    heights = 88
  )
}


# ==============================================================================
# 34. FORMAT DATE COLUMNS
# ==============================================================================
#
# Applies a consistent YYYY-MM-DD format to date fields across the workbook.
#
# ==============================================================================

date_style <- openxlsx::createStyle(
  numFmt = "yyyy-mm-dd"
)

date_columns_by_sheet <- list(
  "Group Sessions" =
    c(
      "Visit Date"
    ),

  "Provider Attribution" =
    c(
      "Visit Date"
    ),

  "Raw Data" =
    c(
      "Visit Date",
      "Canc Date"
    ),

  "Data Quality Detail" =
    c(
      "Visit Date"
    ),

  "Files Processed" =
    c(
      "Earliest Visit Date",
      "Latest Visit Date"
    )
)

purrr::walk(
  names(
    date_columns_by_sheet
  ),
  function(sheet_name) {

    sheet_data <-
      worksheet_data[[sheet_name]]

    if (nrow(sheet_data) == 0) {
      return(invisible(NULL))
    }

    date_column_numbers <- which(
      names(sheet_data) %in%
        date_columns_by_sheet[[sheet_name]]
    )

    if (
      length(
        date_column_numbers
      ) > 0
    ) {

      openxlsx::addStyle(
        wb = workbook,
        sheet = sheet_name,
        style = date_style,
        rows = 2:(nrow(sheet_data) + 1),
        cols = date_column_numbers,
        gridExpand = TRUE,
        stack = TRUE
      )
    }

    invisible(NULL)
  }
)




# ==============================================================================
# 34A. HIDE SUPPORTING CHART WORKSHEETS
# ==============================================================================

openxlsx::sheetVisibility(workbook)[
  which(names(workbook) == "Chart Data")
] <- "hidden"

openxlsx::sheetVisibility(workbook)[
  which(names(workbook) == "Chart Lists")
] <- "hidden"

# ==============================================================================
# 35. SAVE THE VERSION 3 OUTPUT WORKBOOK
# ==============================================================================
#
# Saves a timestamped Version 3 workbook in the output folder. Timestamped
# filenames preserve earlier report versions.
#
# ==============================================================================

output_timestamp <- format(
  Sys.time(),
  "%Y%m%d_%H%M%S"
)

output_file <- file.path(
  output_directory,
  paste0(
    paste0("Clinic_Capacity_", settings$mode, "_"),
    output_timestamp,
    ".xlsx"
  )
)

cc_add_hours_sandbox(workbook, decision_result)
cc_format_decision_sheets(workbook, decision_result)

cc_save_workbook(workbook, output_file)


# ==============================================================================
# 36. REMOVE THE PASSWORD FROM THE R ENVIRONMENT
# ==============================================================================
#
# Removes the shared workbook password from the active R environment after the
# input and output processing has completed.
#
# ==============================================================================

rm(
  workbook_password
)


# ==============================================================================
# 37. DISPLAY COMPLETION DETAILS
# ==============================================================================
#
# Prints the main processing results and the output-file location in the R
# console.
#
# ==============================================================================

exactly_four_count <- sum(
  visits$time_category ==
    "4:00 PM to 4:14 PM",
  na.rm = TRUE
)

four_fifteen_or_later_count <- sum(
  visits$time_category ==
    "4:15 PM or later",
  na.rm = TRUE
)

before_four_count <- nrow(
  visits_before_4_pm
)

all_valid_time_count <- nrow(
  visits_all_time_frames
)

qualifying_group_visit_count <- sum(
  qualifying_visits$potential_group_visit,
  na.rm = TRUE
)

qualifying_group_session_count <-
  dplyr::n_distinct(
    qualifying_visits$session_id[
      qualifying_visits$potential_group_visit
    ],
    na.rm = TRUE
  )

qualifying_completed_individual_count <- sum(
  qualifying_visits$group_visit_classification ==
    "Individual visit" &
    qualifying_visits$manager_status_category ==
    "Completed",
  na.rm = TRUE
)

qualifying_cancelled_individual_count <- sum(
  qualifying_visits$group_visit_classification ==
    "Individual visit" &
    qualifying_visits$manager_status_category ==
    "Cancelled/No-Show",
  na.rm = TRUE
)

qualifying_completed_group_session_count <-
  dplyr::n_distinct(
    qualifying_visits$session_id[
      qualifying_visits$potential_group_visit &
        qualifying_visits$group_session_completed
    ],
    na.rm = TRUE
  )

qualifying_cancelled_group_session_count <-
  dplyr::n_distinct(
    qualifying_visits$session_id[
      qualifying_visits$potential_group_visit &
        qualifying_visits$group_session_cancelled_no_show
    ],
    na.rm = TRUE
  )

message("")
message(
  "=============================================================="
)

message(
  "Clinic Capacity Planner 0.1 - original V3 analysis completed successfully."
)

message(
  "=============================================================="
)

message(
  "Files imported: ",
  length(input_files)
)

message(
  "Total visit rows imported: ",
  format(
    nrow(visits),
    big.mark = ","
  )
)

message(
  "Visits before 4:00 PM: ",
  format(
    before_four_count,
    big.mark = ","
  )
)

message(
  "Visits from 4:00 PM to 4:14 PM: ",
  format(
    exactly_four_count,
    big.mark = ","
  )
)

message(
  "Visits at 4:15 PM or later: ",
  format(
    four_fifteen_or_later_count,
    big.mark = ","
  )
)

message(
  "All visits with a valid appointment time: ",
  format(
    all_valid_time_count,
    big.mark = ","
  )
)

message(
  "Inferred group visits at 4:00 PM or later: ",
  format(
    qualifying_group_visit_count,
    big.mark = ","
  )
)

message(
  "Inferred group sessions at 4:00 PM or later: ",
  format(
    qualifying_group_session_count,
    big.mark = ","
  )
)

message(
  "Completed individual visits at 4:00 PM or later: ",
  format(
    qualifying_completed_individual_count,
    big.mark = ","
  )
)

message(
  "Cancelled/No-Show individual visits at 4:00 PM or later: ",
  format(
    qualifying_cancelled_individual_count,
    big.mark = ","
  )
)

message(
  "Completed group sessions at 4:00 PM or later: ",
  format(
    qualifying_completed_group_session_count,
    big.mark = ","
  )
)

message(
  "Cancelled/No-Show group sessions at 4:00 PM or later: ",
  format(
    qualifying_cancelled_group_session_count,
    big.mark = ","
  )
)

message("")
message(
  "Excel report created at:"
)

message(
  normalizePath(
    output_file,
    winslash = "/",
    mustWork = FALSE
  )
)

message(
  "=============================================================="
)


# ==============================================================================
# END OF CLINIC CAPACITY PLANNER - HOSPITAL REPORT PIPELINE
# ==============================================================================
