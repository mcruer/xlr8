utils::globalVariables(c(
  ".", "actual_col_length", "actual_row_end", "alignment_horizontal",
  "alignment_indent", "alignment_textRotation", "alignment_vertical",
  "alignment_wrapText", "cell_contents", "colNames", "col_name",
  "col_start", "cols", "data", "df", "end_row", "fgFill",
  "fill_patternFill_fgColor_rgb", "fontColour", "font_bold",
  "font_color_rgb", "font_italic", "font_name", "font_size",
  "font_strike", "font_underline", "halign", "sheet_name",
  "style", "rows", "row_start", "col_start", "table_end",
  "table_end_row", "table_length", "textDecoration", "tbl", "value",
  "var", "path", "property", "numFmt", "protection_hidden",
  "protection_locked", "openxlsx_style", "local_format_id", "sheet",
  "name", "row_end", "source_row", "source_col", "target_row", "target_col",
  "formula_to_write", "formula", "formulas", "formula_out", "refs",
  "replacement", "out", "ref", "original", "col_part", "col_abs", "col_idx",
  "row_abs", "row_part", "new_col_idx", "new_row", "new_col", "row_delta",
  "col_delta", "delta_row", "delta_col", "formula_location_raw",
  "formula_from_row", "formula_from_col", "result",
  "column", "dims", "form_metadata", "form_name", "formula_location",
  "key_matches", "keys", "okay", "problem", "raw_df", "reference",
  "row_id", "shifted", "step_1", "step_1_length", "table_end_tbl",
  "tag", "tag_count", "unknown_keys", "x1"
))



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
#' @importFrom openxlsx2 int2col col2int
NULL




rowcol_to_ref <- function(row, col, col_abs = FALSE, row_abs = FALSE) {
  tibble(row, col, col_abs, row_abs) %>%
    mutate(
      ref = paste0(
        if_else(col_abs, "$", ""),
        int2col(col),
        if_else(row_abs, "$", ""),
        row
      )
    ) %>%
    pull(ref)
}

ref_to_row <- function(refs) {
  as.integer(stringr::str_extract(refs, "\\d+$"))
}

ref_to_col <- function(refs) {
  col2int(gsub("\\$?[0-9]+", "", refs))
}

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


#' Shift Cell References in Excel Formulas
#'
#' Given an Excel formula and a change in location (source → target cell), this function
#' adjusts all relative cell references to reflect the new position, while preserving absolute references.
#'
#' This is useful when copying formulas between cells and wanting to simulate Excel's built-in reference shifting logic.
#' It handles single-cell references, ranges (e.g., \code{A1:B2}), and complex expressions (e.g., \code{=IF(A1>0, B1, C1)}).
#'
#' @param formula Character vector of Excel formulas.
#' @param source_row Integer vector of row positions where each formula originally resided.
#' @param source_col Integer vector of column positions where each formula originally resided.
#' @param target_row Integer vector of row positions where the formula is being moved to.
#' @param target_col Integer vector of column positions where the formula is being moved to.
#'
#' @return A character vector of formulas with shifted cell references, aligned with their new positions.
#'
#' @examples
#' shift_formula_references("=A$1+$B2", source_row = 1, source_col = 1, target_row = 3, target_col = 3)
#' # Returns: "=C$1+$B4"
#'
#' @importFrom tibble tibble
#' @importFrom tidyr unnest
#' @importFrom dplyr mutate select pull if_else
#' @importFrom purrr pmap pmap_chr map map2
#' @importFrom stringr str_extract str_extract_all str_replace_all fixed
#' @importFrom openxlsx2 col2int int2col
#' @export
shift_formula_references <- function(formula, source_row, source_col, target_row, target_col) {

  shift_refs <- function(refs, row_delta, col_delta) {

    row_delta <- row_delta[1]
    col_delta <- col_delta[1]

    tibble(original = refs) %>%
      unnest(original) %>%
      mutate(
        col_abs = grepl("^\\$", original),
        row_abs = grepl("\\$", str_extract(original, "\\$?[0-9]+")),
        col_part = gsub("\\$.*", "", gsub("^\\$", "", original)),
        row_part = as.integer(str_extract(original, "[0-9]+")),
        col_idx = col2int(gsub("\\$", "", col_part)),
        new_col_idx = if_else(col_abs, col_idx, col_idx + col_delta),
        new_row = if_else(row_abs, row_part, row_part + row_delta),
        new_col = int2col(new_col_idx),
        shifted = paste0(
          if_else(col_abs, "$", ""),
          new_col,
          if_else(row_abs, "$", ""),
          new_row
        )
      )
  }

  replace_formula_literals <- function(string, pattern, replacement) {
    stopifnot(length(pattern) == length(replacement))
    for (i in seq_along(pattern)) {
      string <- str_replace_all(string, fixed(pattern[i]), replacement[i])
    }
    string
  }

  tibble(formula, source_row, source_col, target_row, target_col) %>%
    mutate(
      row_delta = target_row - source_row,
      col_delta = target_col - source_col,
      refs = str_extract_all(formula, "\\$?[A-Z]{1,3}\\$?[0-9]{1,7}"),
      shifted = pmap(list(refs, row_delta, col_delta), shift_refs),
      replacement = map(shifted, pull, shifted),
      out = pmap_chr(list(formula, refs, replacement), replace_formula_literals)
    ) %>%
    pull(out)
}
