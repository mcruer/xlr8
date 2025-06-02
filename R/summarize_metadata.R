#' @importFrom dplyr pull filter first
#' @importFrom gplyr filter_out_na
extract_form_name <- function(raw_df) {
  out <- raw_df %>%
    filter(sheet_name == "form") %>%
    filter_out_na(x1) %>%
    pull(x1) %>%
    first()

  if (is.null(out)) {
    return(NA_character_)
  }

  return(out)
}


#' Extract Specific Tag Key from Metadata Tag
#'
#' Parses a string to extract a specific metadata key (e.g., "var", "tbl", "col", etc.)
#' from an xlr8-style metadata tag.
#'
#' @param string A character vector of tag strings.
#' @param target A character string indicating which tag key to extract (e.g., "var").
#'
#' @return A character vector containing the extracted key values or NA if not present.
#'
#' @importFrom stringr str_c str_remove
#' @importFrom gplyr replace_with_na
#'
#' @keywords internal
target_extraction <- function(string, target) {
  front <- str_c(".*", xlr8_tag, target, xlr8_tag)
  str_remove(string, front) %>%
    str_remove(str_c(xlr8_tag, ".*")) %>%
    replace_with_na("^$")
}

#' Summarize Metadata Tags from Parsed Excel Template
#'
#' Processes an Excel-derived tibble and extracts structured metadata from embedded tags.
#' Handles variables, tables, columns, and formulas. Also computes inferred dimensions.
#'
#' @param raw_template A tibble from `read_excel_all()` representing cell content across Excel sheets.
#'
#' @return A tibble with one row and five list-columns:
#' \describe{
#'   \item{tags_info}{Raw tag positions.}
#'   \item{variable_info}{Metadata about single-cell variables.}
#'   \item{table_info}{Table extents and positions.}
#'   \item{column_info}{Column names and positions within tables.}
#'   \item{all_info}{Unified metadata for writing and styling workflows.}
#' }
#'
#' @importFrom dplyr mutate select filter arrange group_by ungroup rename relocate summarise bind_rows if_else
#' @importFrom tidyr pivot_longer fill
#' @importFrom stringr str_c str_remove
#' @importFrom purrr map_chr map
#' @importFrom gplyr replace_with_na filter_out_na
#' @importFrom tibble tibble
#'
#' @export
summarize_metadata_from_raw_template <- function(raw_template) {

  form <- extract_form_name(raw_template)

  resolve_formula_reference <- function(formula_location, row, col) {
    tibble(formula_location, row, col) %>%
      mutate(
        is_abs = grepl("^[A-Z]+[0-9]+$", toupper(formula_location)),
        delta_row = case_when(
          tolower(formula_location) == "up" ~ -1,
          tolower(formula_location) == "down" ~ 1,
          is.na(formula_location) ~ NA,
          TRUE ~ 0
        ),
        delta_col = case_when(
          tolower(formula_location) == "left" ~ -1,
          tolower(formula_location) == "right" ~ 1,
          TRUE ~ 0
        ),
        new_row = row + delta_row,
        new_col = col + delta_col,
        result = case_when(
          is_abs ~ toupper(formula_location),
          is.na(new_row) | is.na(new_col) | new_row < 1 | new_col < 1 ~ NA_character_,
          TRUE ~ map2_chr(new_row, new_col, ~ paste0(openxlsx2::int2col(.y), .x))
        )
      ) %>%
      pull(result)
  }
  # Extract tags and parse components
  full_tags <- raw_template %>%
    pivot_longer(-c(row, sheet_name)) %>%
    quickm(name, str_remove, "x") %>%
    to_number(name) %>%
    rename(col = name) %>%
    relocate(row, col, .after = sheet_name) %>%
    filter_in(value, str_c("^", xlr8_tag), na.rm = TRUE) %>%
    mutate(
      var = target_extraction(value, "var"),
      tbl = target_extraction(value, "tbl"),
      col_name = target_extraction(value, "col"),
      table_end_tbl = target_extraction(value, "table_end"),
      formula_location_raw = target_extraction(value, "formula"),
      formula_location = resolve_formula_reference(formula_location_raw, row, col)
    )

  # Save raw tags as audit log
  tags_info <- full_tags %>%
    select(sheet_name, row, col, value)

  # Variable metadata
  variable_info <- full_tags %>%
    filter(!is.na(var)) %>%
    select(sheet_name, row, col, var, formula_location)

  end_rows <- full_tags %>%
    filter_out_na(table_end_tbl) %>%
    select(sheet_name, tbl = table_end_tbl, end_row = row)

  table_lengths <- full_tags %>%
    filter_out_na(tbl) %>%
    group_by(sheet_name, tbl) %>%
    summarise(.groups = "drop") %>%
    left_join(end_rows, by = c("sheet_name", "tbl"))

  # Extract columns within tables
  table_info_interim <- full_tags %>%
    filter(!is.na(col_name)) %>%
    arrange(sheet_name, row, col) %>%
    group_by(sheet_name) %>%                         # 👈 FIXED SCOPE
    fill(tbl, .direction = "down") %>%
    ungroup() %>%
    select(sheet_name, row, col, tbl, col_name, formula_location) %>%
    left_join(table_lengths, by = c("sheet_name", "tbl")) %>%
    group_by(sheet_name, tbl)

  # Handle templates with no tables
  if (nrow(table_info_interim) == 0) {
    table_info <- tibble(
      sheet_name = factor(),
      tbl = character(),
      row_start = double(),
      row_end = double(),
      col_start = double(),
      col_end = double()
    )

    column_info <- tibble(
      sheet_name = factor(),
      col = double(),
      tbl = character(),
      col_name = character(),
      formula_location = character()
    )
  } else {
    table_info <- table_info_interim %>%
      fill(end_row, .direction = "down") %>%
      summarise(
        row_start = min(.data$row),
        row_end = mean(.data$end_row),
        col_start = min(.data$col),
        col_end = max(.data$col)
      ) %>%
      ungroup()

    column_info <- full_tags %>%
      filter(!is.na(col_name)) %>%
      arrange(sheet_name, row, col) %>%
      group_by(sheet_name) %>%                       # 👈 FIXED SCOPE
      fill(tbl, .direction = "down") %>%
      ungroup() %>%
      group_by(sheet_name, tbl) %>%
      mutate(col = col - min(col) + 1) %>%
      select(sheet_name, col, tbl, col_name, formula_location) %>%
      ungroup()
  }

  # Reformat variable_info to match table_info structure
  variable_info_reformed <- variable_info %>%
    rename(tbl = var,
           row_start = row,
           col_start = col) %>%
    mutate(row_end = row_start,
           col_end = col_start)

  # Construct all_info metadata
  all_info <- table_info %>%
    left_join(column_info, by = c("sheet_name", "tbl")) %>%
    mutate(
      col = col - 1,
      col_start = col_start + col
    ) %>%
    select(-col) %>%
    bind_rows(variable_info_reformed) %>%
    mutate(
      form = form,
      row_id = if_else(is.na(col_name),
                       str_c(form, tbl, sep = " "),
                       str_c(form, tbl, col_name, sep = " "))
    ) %>%
    relocate(row_id, .before = 1)

  #Add functions to write.

  formula_definitions <- gplyr::uncloak(raw_template)$formulas %>%
    mutate(formula_location = rowcol_to_ref(row, col)) %>%
    select(sheet_name, formula, formula_location)

  output_formulas<- all_info %>%
    filter_out_na(formula_location) %>%
    left_join(formula_definitions)%>%
    mutate(
      formula_from_row = ref_to_row(formula_location),
      formula_from_col = ref_to_col(formula_location),
      formula_to_write = shift_formula_references(formula,
                                                  formula_from_row,
                                                  formula_from_col,
                                                  row_start, col_start
      ))  %>%
    select(formula_location, formula_to_write)

  all_info <- all_info %>%
    left_join(output_formulas) %>%
    relocate(formula_to_write, .after = formula_location)


  # Return all metadata
  tibble(
    tags_info = list(tags_info),
    variable_info = list(variable_info),
    table_info = list(table_info),
    column_info = list(column_info),
    all_info = list(all_info)
  )
}

#' Validate Raw Tag Strings for Unbalanced or Unknown Keys
#'
#' Checks if metadata tag strings have an uneven number of openers or contain unknown tag keys.
#'
#' @param tag_values A character vector of metadata tag strings (e.g., from a raw Excel tibble).
#' @param allowed_keys Character vector of valid tag keys (defaults to standard xlr8 tags).
#'
#' @return A list with two tibbles:
#' \describe{
#'   \item{unbalanced}{Rows with an uneven number of opening tags.}
#'   \item{unknown}{Rows with unknown tag keys.}
#' }
#'
#' @importFrom dplyr mutate filter
#' @importFrom stringr str_count str_match_all fixed
#' @importFrom purrr map_chr map
#' @importFrom tibble tibble
#'
#' @export
validate_metadata_tags <- function(tag_values, allowed_keys = c("var", "tbl", "col", "formula", "table_end")) {
  tag_df <- tibble(value = tag_values)

  tag_df <- tag_df %>%
    mutate(
      tag_count = str_count(value, fixed("*((")),
      key_matches = str_match_all(value, "\\*\\(\\(([^*]+)\\*\\(\\("),
      keys = map(key_matches, ~ .x[,2]),
      unknown_keys = map_chr(keys, ~ paste(setdiff(.x, allowed_keys), collapse = ", "))
    )

  list(
    unbalanced = tag_df %>% filter(tag_count %% 2 != 0),
    unknown = tag_df %>% filter(unknown_keys != "")
  )
}


#' Validate Parsed Metadata from Excel Template
#'
#' Runs a series of checks on extracted metadata to catch issues like unbalanced tags,
#' whitespace, missing table definitions, or duplicate tags.
#'
#' @param metadata A list-like object output from `summarize_metadata_from_raw_template()`.
#' @param quiet Logical; suppress warning output and printing if TRUE.
#' @param fail_on_issue Logical; stop execution if problems are found (default TRUE).
#'
#' @return Returns metadata invisibly if no issues are found, otherwise stops with an error.
#'
#' @importFrom dplyr mutate select rename relocate bind_rows filter full_join ungroup rowwise if_else
#' @importFrom tidyr unnest
#' @importFrom purrr map_chr map2
#' @importFrom stringr str_count fixed
#' @importFrom gplyr filter_in filter_out_na filter_in_na
#' @importFrom openxlsx2 wb_dims
#' @importFrom janitor get_dupes
#' @importFrom gplyr filter_out
#' @export
validate_metadata <- function(metadata, quiet = FALSE, fail_on_issue = TRUE) {

  variable_info <- pull_cell(metadata, variable_info)
  column_info <- pull_cell(metadata, column_info)
  tags_info <- pull_cell(metadata, tags_info)
  table_info <- pull_cell(metadata, table_info)

  all_info <- pull_cell(metadata, all_info)

  check <- all_info %>%
    full_join(tags_info %>%
                rename(
                  tag = value,
                  row_start = row,
                  col_start = col
                )) %>%
    relocate(tag) %>%
    select(-row_id) %>%
    suppressMessages()

  ghost_table_ends <- tags_info %>%
    filter_in(value, "\\*\\(\\(table_end") %>%
    mutate(table = target_extraction(value, "table_end")) %>%
    left_join(table_info %>% select(table = tbl) %>% mutate(okay = TRUE)) %>%
    filter_in_na(okay) %>%
    select(-okay) %>%
    rename(tag = value, row_start = row, col_start = col, tbl = table) %>%
    mutate(problem = "Table End Without a Matching Table") %>%
    suppressMessages()


  white_space <- check %>%
    filter_in(tag, "\\s", na.rm = TRUE) %>%
    mutate(problem = "Whitespace in Tag")

  unballanced_tags <- tags_info  %>%
    mutate(tag_count = str_count(value, fixed("*(("))) %>%
    filter(tag_count %% 2 != 0) %>%
    mutate(problem = "Uneven Number of Tags - Suggests Something was Missed") %>%
    rename(tag = value, row_start = row, col_start = col)

  duplicate_tags <- check %>%
    filter_out_na(tag) %>%
    get_dupes(tag, tbl) %>%
    mutate(problem = "Duplicate Tag") %>%
    suppressMessages()

  other_errors <- check %>%
    filter_in_na(tag, tbl, if_any_or_all = "if_any") %>%
    mutate(problem = "Some Other Issue")

  all_errors <- bind_rows(duplicate_tags, other_errors) %>%
    bind_rows(white_space) %>%
    bind_rows(unballanced_tags) %>%
    filter_out(tag, "\\*\\(\\(table_end") %>%
    bind_rows(ghost_table_ends) %>%
    rowwise() %>%
    mutate(reference = wb_dims(row_start, col_start), .after = sheet_name
    ) %>%
    ungroup() %>%
    select(tag,
           sheet_name,
           reference,
           tbl, col_name,
           problem) %>%
    suppressWarnings()

  if (nrow(all_errors) > 0 && !quiet) {
    warning("Metadata validation found issues. See printed tibble for details.")
    print(all_errors)
  }

  if (fail_on_issue && nrow(all_errors) > 0) {
    stop("Metadata validation failed. See tibble for errors.")
  }

  invisible(metadata)
}


#' Summarize Metadata Directly from an Excel Template File
#'
#' Reads an Excel template file and summarizes structured metadata tags embedded in the sheets,
#' extracting comprehensive details about variables, tables, and column definitions.
#' This function is a convenient wrapper around \code{read_excel_all()} followed by internal metadata parsing.
#'
#' @param metadata_path File path to the Excel template workbook.
#' @param sheets Optional character vector specifying exact sheet names to read.
#'   Defaults to \code{NULL}, in which case sheets matching \code{sheets_regex} are read.
#' @param sheets_regex Regular expression pattern to select sheets when \code{sheets} is \code{NULL}.
#'   Defaults to \code{"."}, matching all sheets.
#' @param validate Logical. Whether to validate parsed metadata for issues such as unbalanced tags or duplicates.
#'   Defaults to \code{TRUE}.
#' @param quiet Logical. If \code{TRUE}, suppresses printed warnings and messages during validation.
#'   Defaults to \code{FALSE}.
#' @param fail_on_issue Logical. If \code{TRUE}, causes validation to stop execution when problems are found.
#'   Defaults to \code{TRUE}.
#'
#' @return A tibble containing structured metadata organized into four list-columns:
#'   \itemize{
#'     \item \code{variable_info}: Metadata describing individual cells (variables), including:
#'       \itemize{
#'         \item \code{sheet_name}: Sheet where variable is located.
#'         \item \code{row_start}, \code{col_start}: Position (row and column) of the variable.
#'         \item \code{tbl}: Name assigned to the variable.
#'       }
#'
#'     \item \code{table_info}: Metadata summarizing each table identified in the Excel template, including:
#'       \itemize{
#'         \item \code{sheet_name}: Sheet name containing the table.
#'         \item \code{tbl}: Table name extracted from metadata tags.
#'         \item \code{row_start}, \code{row_end}: Row range of the table.
#'         \item \code{col_start}, \code{col_end}: Column range of the table.
#'       }
#'
#'     \item \code{column_info}: Metadata detailing columns within each table, including:
#'       \itemize{
#'         \item \code{sheet_name}: Sheet name containing the column.
#'         \item \code{tbl}: Associated table name.
#'         \item \code{col}: Column position relative to its table (starting from 1).
#'         \item \code{col_name}: Name assigned to the column.
#'       }
#'
#'     \item \code{all_info}: Combined metadata tibble containing both variable and column information
#'       suitable for subsequent data-writing and styling functions.
#'   }
#'
#' @details
#' The Excel template must have metadata annotations embedded using a predefined format (\code{xlr8_tag}),
#' which enables the structured identification of variables, tables, and columns. Specifically, this function:
#'
#' \itemize{
#'   \item Extracts metadata from single-cell variables indicated by tags like \code{*((var*((...))}.
#'   \item Identifies tables by start/end markers like \code{*((tbl*((...))} and \code{*((table_end}.
#'   \item Extracts column definitions within tables using \code{*((col*((...))} tags.
#' }
#'
#' The function performs extensive internal checks for errors such as misaligned columns or incomplete metadata.
#' If misalignments or other issues are detected, the function stops execution and provides clear, explicit
#' error messages to facilitate correction.
#'
#' Users who require separate control over reading Excel sheets and summarizing metadata
#' can independently use \code{read_excel_all()} and internal summarization functions.
#'
#' @examples
#' \dontrun{
#' metadata <- summarize_metadata("template.xlsx")
#' metadata$table_info  # View detailed table metadata
#' metadata$variable_info  # Inspect extracted variable metadata
#' }
#'
#' @export
summarize_metadata <- function(metadata_path, sheets = NULL, sheets_regex = ".",
                               validate = TRUE, quiet = FALSE, fail_on_issue = TRUE) {
  raw_template <- read_excel_all(metadata_path, sheets, sheets_regex)
  metadata <- summarize_metadata_from_raw_template(raw_template)

  if (validate) {
    validate_metadata(metadata, quiet = quiet, fail_on_issue = fail_on_issue)
  }

  return(metadata)
}
