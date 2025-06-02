#' Write Formulas to an Excel Workbook
#'
#' Inserts formulas into an Excel workbook using metadata-defined positions. This function
#' shifts formulas as needed (e.g., for column-based replication) and writes them using
#' `openxlsx2::wb_add_formula()`.
#'
#' It assumes that the input metadata has been enriched with a `formula_to_write` column,
#' typically via `summarize_metadata()` which processes embedded Excel template tags.
#'
#' @param wb A `wbWorkbook` object loaded with `openxlsx2::wb_load()`.
#' @param df A data frame of structured input data, where list-columns contain table-like tibbles.
#' @param all_info A tibble of metadata (from `summarize_metadata()`) that includes
#'   positioning, structure, and `formula_to_write` values.
#'
#' @return The modified workbook object with formulas written in place.
#'
#' @examples
#' \dontrun{
#' wb <- openxlsx2::wb_load("template.xlsx")
#' metadata <- summarize_metadata("template.xlsx")
#' df <- tibble::tibble(
#'   summary_table = list(tibble::tibble(score = 1:3))
#' )
#' wb <- add_formulas(wb, df, metadata$all_info[[1]])
#' openxlsx2::wb_save(wb, "with_formulas.xlsx", overwrite = TRUE)
#' }
#'
#' @importFrom dplyr select where mutate left_join if_else rename
#' @importFrom purrr map_int pmap pmap_chr
#' @importFrom tidyr pivot_longer unnest
#' @importFrom listful pbuild
#' @importFrom openxlsx2 wb_add_formula
#' @importFrom gplyr quickm filter_out_na
#' @importFrom tibble tibble
#' @export
add_formulas <- function (wb, df, all_info) {
  list_formulas <- function(formula, row_start, col_start, row_end) {
    tibble(
      formula = formula,
      target_row = row_start:row_end,
      target_col = col_start,
      source_row = row_start,
      source_col = col_start
    ) %>%
      mutate(formula_out = shift_formula_references(formula, source_row, source_col, target_row, target_col))
  }


  table_lengths <- df %>%
    select(where(is.list)) %>%
    pivot_longer(cols = everything(),
                 names_to = "tbl",
                 values_to = "table_length") %>%
    quickm(table_length, ~ map_int(.x, nrow))

  formula_map <- all_info %>%
    left_join(table_lengths, by = "tbl") %>%
    mutate(row_end = if_else(is.na(row_end), row_start + table_length - 1, row_end)) %>%
    select(sheet = sheet_name, formula_to_write, row_start, col_start, row_end) %>%
    filter_out_na(formula_to_write) %>%
    rename(formula = formula_to_write) %>%
    mutate(formulas = pmap(list(formula, row_start, col_start, row_end), list_formulas)) %>%
    select(sheet, formulas) %>%
    unnest(formulas) %>%
    select(sheet,
           start_row = target_row,
           start_col = target_col,
           x = formula_out)

  pbuild(wb, formula_map, wb_add_formula)

}
