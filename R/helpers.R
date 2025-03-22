#' @importFrom dplyr select mutate filter rename left_join bind_rows relocate ungroup case_when
#' @importFrom tidyr pivot_longer pivot_wider
#' @importFrom purrr map map2 map_int pmap pwalk
#' @importFrom stringr str_c str_sub str_extract str_remove str_remove_all str_replace_all
#' @importFrom tibble as_tibble tibble
#' @importFrom readr parse_guess
#' @importFrom openxlsx loadWorkbook saveWorkbook createStyle addStyle
#' @importFrom utils capture.output head
#' @importFrom magrittr %>%
#' @importFrom rlang .data
#' @importFrom gplyr quickm filter_out_na filter_in_na to_number pull_cell replace_with_na quicks
#' @importFrom gplyr filter_in parse_guess_all
#' @importFrom tidyxl xlsx_cells xlsx_formats
#' @importFrom dplyr rowwise
#' @importFrom stats na.omit
NULL

#' Metadata Tag Pattern
#'
#' Regex pattern used to detect metadata tags in Excel sheets.
#' Primarily for internal use.
#'
#' @noRd
xlr8_tag <- "\\*\\(\\("

#' Get Top Rows of a Data Frame
#'
#' Returns the first `table_length` rows of `data`. If `table_length` is `NA`,
#' returns the entire data frame.
#'
#' @param data Data frame from which rows will be extracted.
#' @param table_length Integer specifying the number of top rows to keep, or `NA` to keep all rows.
#'
#' @return A data frame containing the specified number of top rows.
#'
#' @examples
#' get_top_of_table(mtcars, 5)
#' get_top_of_table(mtcars, NA)
#'
#' @noRd
get_top_of_table <- function (data, table_length) {
  if (is.na(table_length)) {
    return(data)
  }
  return(head(data, table_length))
}

#' Prepare Column-Based Metadata and Data
#'
#' Extracts and organizes data for Excel columns based on provided metadata.
#' Filters metadata for entries that have column names defined, and
#' constructs data frames corresponding to each table column as specified.
#'
#' @param all_info Tibble containing metadata information about tables, columns, and positions.
#' @param df Data frame from which data is extracted.
#'
#' @return Tibble containing metadata alongside extracted column-specific data,
#'         ready for writing to Excel.
#'
#' @details
#' This function internally filters out entries where column names (`col_name`) are missing (`NA`).
#' It then extracts specified columns from the provided data frame (`df`) based on table (`tbl`)
#' and column name (`col_name`) information found in `all_info`. It also calculates
#' the actual length of each extracted table segment.
#'
#' @noRd
all_info_cols <- function (all_info, df) {
  all_info %>%
    filter_out_na(col_name) %>%
    mutate(
      df = list(df),
      data = map2(df, tbl, ~ .x %>%
                    pull_cell(.y)) ,
      data =
        map2(data, col_name, ~ .x %>% select(all_of(.y))),
      table_length = row_end - row_start + 1,
      data = map2(data,
                  table_length,
                  get_top_of_table),
      colNames = FALSE,
      actual_col_length = map_int(data, nrow),
      actual_row_end = row_start +actual_col_length - 1,
    )
}


#' Prepare Variable-Based Metadata and Data
#'
#' Extracts data for single-cell variables (non-column data) based on provided metadata.
#' Filters metadata entries without column names, indicating individual cells rather than entire columns.
#'
#' @param all_info Tibble containing metadata information about variables, tables, and positions.
#' @param df Data frame from which individual variable data is extracted.
#'
#' @return Tibble containing metadata alongside extracted single-cell data,
#'         ready for writing to Excel.
#'
#' @details
#' This function internally filters metadata to include only rows where `col_name` is `NA`.
#' It then extracts individual cell values from the provided data frame (`df`) based on table (`tbl`)
#' information found in `all_info`.
#'
#' @noRd
all_info_vars <- function (all_info, df) {
  all_info %>%
    filter_in_na(col_name) %>%
    mutate(
      df = list(df),
      data = map2(df, tbl, ~ .x %>%
                    pull_cell(.y)),
      colNames = FALSE,
    )
}
