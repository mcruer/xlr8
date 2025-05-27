#' Apply Excel Styles Using Reference Cell Copying (openxlsx2)
#'
#' Applies Excel styles by copying formatting from template reference cells to populated data ranges.
#' This method uses native style inheritance via `wb_set_cell_style()` and assumes the workbook was
#' created using `openxlsx2`.
#'
#' @param wb A `wbWorkbook` object created with `openxlsx2::wb_load()` or `wb_workbook()`.
#' @param all_info A tibble from `summarize_metadata()` that includes layout and tagging info.
#' @param df A tibble where some columns are list-columns containing tables (e.g., from `xlr8_read()` or manually constructed).
#'
#' @return A workbook object with styles applied based on template reference cells.
#' @importFrom dplyr where
#' @importFrom purrr pmap_chr map2_chr
#' @importFrom openxlsx2 wb_dims wb_set_cell_style
#' @importFrom listful pbuild
#'
#' @export
apply_styles <- function(wb, df, all_info) {
  table_lengths <- df %>%
    select(where(is.list)) %>%
    pivot_longer(cols = everything(), names_to = "tbl", values_to = "table_length") %>%
    quickm(table_length, ~ map_int(.x, nrow))

  style_map <- all_info %>%
    left_join(table_lengths, by = "tbl") %>%
    mutate(
      row_end = if_else(is.na(row_end), row_start + table_length - 1, row_end),
      dims = pmap_chr(list(row_start, row_end, col_start), ~ wb_dims(..1:..2, ..3)),
      style = map2_chr(row_start, col_start, wb_dims)
    ) %>%
    select(sheet = sheet_name, dims, style)

  pbuild(wb, style_map, wb_set_cell_style)
}
