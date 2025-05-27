#' Read Structured Data from Tagged Excel Forms
#'
#' Extracts structured data from Excel workbooks that have been pre-tagged with metadata
#' using the `xlr8` system. This function reads both single-cell variables and multi-row
#' tables, guided by positional metadata produced by `summarize_metadata()`.
#'
#' @param raw_df A tibble representing the contents of an Excel workbook, typically
#'   produced by `read_excel_all()`. It must include sheet names, row numbers, and one
#'   column per Excel column (e.g., `x1`, `x2`, etc.).
#' @param all_info A tibble of metadata produced by `summarize_metadata()` (specifically
#'   the `all_info` entry). It defines the position of each variable and table/column.
#' @param fix_dates_regex Optional string specifying a regular expression used to detect
#'   column names that should be converted from Excel numeric date format. Defaults to
#'   `"date_|_date|^date$"`. Set to `NULL` to skip date conversion.
#'
#' @return A tibble with a single row:
#'   - Each variable becomes a column with a single value.
#'   - Each table becomes a nested tibble-column, with its name derived from the metadata.
#'
#' @details
#' This function assumes the Excel input was generated from a locked, tagged template.
#' The metadata must indicate the location of each variable and table using `*((var((...))`,
#' `*((tbl((...))`, and `*((col((...))` tags. It reconstructs tables by matching each
#' column to its metadata-defined position and name, then combines them row-wise.
#'
#' If date columns in tables or variables are stored as numeric Excel dates,
#' they are automatically converted based on a regex match unless `fix_dates_regex` is `NULL`.
#'
#' @examples
#' \dontrun{
#' raw_df <- read_excel_all("form_filled.xlsx")
#' metadata <- summarize_metadata("template.xlsx")
#' xlr8_read(raw_df, metadata$all_info[[1]])
#' }
#'
#' @importFrom dplyr mutate select filter if_else bind_cols across everything matches where
#' @importFrom tidyr pivot_wider nest unnest
#' @importFrom purrr map map2 map_int pmap
#' @importFrom gplyr quickm filter_out_na filter_in_na parse_guess_all
#' @importFrom janitor excel_numeric_to_date
#' @importFrom magrittr extract
#'
#' @export

xlr8_read <- function(raw_df, all_info, fix_dates_regex = "date_|_date|^date$") {

  pull_table <- function(df, top, bottom, left, right){
    df %>%
      extract(top:bottom, left:right)
  }

  data_by_table_and_column_name <- all_info %>%
    mutate(
      raw_df = list(raw_df),
      step_1 = map2(
        raw_df,
        sheet_name,
        ~ .x %>% filter(sheet_name == .y) %>% select(-sheet_name, -row)
      ),
      step_1_length = map_int(step_1, nrow),
      row_end = if_else(is.na(row_end), step_1_length, row_end),
      column = pmap(
        list(
          df = step_1,
          top = row_start,
          bottom = row_end,
          left = col_start,
          right = col_start
        ),
        pull_table
      )
    ) %>%
    select(tbl, col_name, column)

  collapse_into_table <- function (df) {
    df %>%
      pivot_wider(names_from = "col_name", values_from = "column") %>%
      mutate(across(everything(),
                    .fns = ~ map2(.x, cur_column(), set_names))) %>%
      unnest(everything())
  }

  tables <- data_by_table_and_column_name %>%
    filter_out_na(col_name) %>%
    nest(data = -tbl) %>%
    pivot_wider(names_from = "tbl", values_from = "data") %>%
    quickm(everything(), map, collapse_into_table) %>%
    quickm(everything(), map, parse_guess_all)

  fix_dates_function <- function(df) {
    df %>%
      quickm(matches(fix_dates_regex) &
               where(is.numeric),
             excel_numeric_to_date)
  }


  if (!is.null(fix_dates_regex)) {
    tables <- tables %>%
      quickm(everything(), map, fix_dates_function)
  }


  vars <- data_by_table_and_column_name %>%
    filter_in_na(col_name) %>%
    select(-col_name) %>%
    quickm(column, map, unlist) %>%
    pivot_wider(names_from = "tbl", values_from = "column") %>%
    unnest(everything()) %>%
    parse_guess_all()

  if (!is.null(fix_dates_regex)) {
    vars <- vars %>%
      fix_dates_function()
  }

  bind_cols(vars, tables)
}
