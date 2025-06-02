
#' Read All Cells from Excel Sheets into a Structured Tibble
#'
#' Reads and extracts cell data from specified Excel sheets using \code{tidyxl},
#' converting it into a structured, wide-format tibble suitable for analysis.
#' Allows users to select sheets explicitly or using regex patterns.
#'
#' @param path File path to the Excel workbook.
#' @param sheets Optional character vector specifying exact sheet names to read.
#'   Defaults to \code{NULL}, in which case sheets are matched using \code{sheets_regex}.
#' @param sheets_regex Regular expression pattern to select sheets when \code{sheets} is \code{NULL}.
#'   Defaults to \code{"."}, matching all sheets.
#'
#' @return A wide-format tibble structured as follows:
#'   - Each **row** represents a single cell position within a sheet.
#'   - **Columns** include:
#'     - \code{sheet_name} (factor): Name of the Excel sheet (ordered as in workbook).
#'     - \code{row} (integer): Row number from the Excel sheet.
#'     - Columns named \code{x1}, \code{x2}, \code{x3}, ..., each representing Excel columns
#'       by their numerical index. Each cell contains the original Excel cell's content,
#'       or \code{NA} if the cell was empty.
#'
#'   For example, a cell originally located at sheet "Data", row 3, column B in Excel
#'   will appear in the resulting tibble in the row where:
#'   - \code{sheet_name = "Data"}
#'   - \code{row = 3}
#'   - \code{x2} contains the cell's contents.
#'
#' @details
#' The function processes Excel sheets selected by name or regex patterns, extracting
#' non-empty cells using \code{tidyxl::xlsx_cells()}. It handles various Excel data types:
#' character, numeric, blank, and errors. The output tibble is reshaped into a convenient
#' wide-format structure, with each Excel sheet represented clearly and uniformly.
#'
#' @examples
#' \dontrun{
#' # Read all sheets in an Excel workbook
#' read_excel_all("workbook.xlsx")
#'
#' # Read specific sheets by exact names
#' read_excel_all("workbook.xlsx", sheets = c("Data", "Summary"))
#'
#' # Read sheets matching a pattern (e.g., sheets starting with "2024")
#' read_excel_all("workbook.xlsx", sheets_regex = "^2024")
#' }
#'
#' @export
read_excel_all <- function(path, sheets = NULL, sheets_regex = ".") {

  sheet_names <- tidyxl::xlsx_sheet_names(path)
  if(is.null(sheets)){
    sheet_names <- sheet_names %>%
      gplyr::str_filter(sheets_regex)

    if(length(sheet_names)== 0) {
      stop("The sheets_regex argument didn't match any sheets in the workbook.")
    }

  } else {
    problems <- setdiff(sheets, sheet_names)
    if(length(problems) > 0){
      stop(str_c("The following sheets are not in the workbook: ",
                 str_c (problems, sep = ", ", collapse = ", ")))
    }

    sheet_names <- sheets

  }

  initial <- tidyxl::xlsx_cells(path, sheets = sheet_names,
                                include_blank_cells = FALSE,
  ) %>%
    dplyr::filter(sheet %in% sheet_names) %>%
    dplyr::mutate(
      cell_contents = dplyr::case_when(
        data_type == "blank" ~ NA_character_,
        data_type == "character" ~ character,
        data_type == "error" ~ error,
        data_type == "logical" ~ as.character (logical),
        data_type == "date" ~ as.character (date),
        TRUE ~ content,
      )
    ) %>%
    dplyr::select(sheet_name = sheet, row, col, cell_contents, formula)

  out <- initial %>%
    dplyr::select(-formula) %>%
    dplyr::group_by(sheet_name) %>%
    gplyr::quicks(c(row, col), max) %>%
    dplyr::mutate(df = map2(row, col,
                            ~ tidyr::expand_grid(row = 1:.x, col = 1:.y))) %>%
    dplyr::select(-row, -col) %>%
    tidyr::unnest(df) %>%
    dplyr::full_join(initial %>%
                       dplyr::select(-formula)) %>%
    tidyr::pivot_wider(names_from = col, values_from = cell_contents) %>%
    dplyr::rename_with( ~ stringr::str_c("x", .x), -c(sheet_name, row)) %>%
    dplyr::mutate(sheet_name = as.character(sheet_name)) %>%
    dplyr::arrange(sheet_name) %>%
    suppressMessages()

  gplyr::cloak(out, list(formulas = initial, wb = openxlsx2::wb_load (path)))

}
