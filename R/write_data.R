#' Write Data to Excel Workbook Based on Metadata Information
#'
#' Writes data from an R data frame into an Excel workbook using the
#' `openxlsx2` package. It populates the workbook with both table-based data
#' (multiple rows and columns) and single-cell variable data, based on positions
#' and structures defined in the provided metadata.
#'
#' @param wb An Excel workbook object created by \code{openxlsx2::wb_load()}
#'   into which data will be written.
#' @param df Data frame containing the actual data to populate into Excel sheets.
#' @param all_info Tibble containing metadata describing precisely where data
#'   should be placed within the workbook. Typically created by \code{summarize_metadata()}.
#'
#' @return The modified Excel workbook object (\code{wb}) with data written
#'   according to the provided metadata. Note: This function modifies the workbook
#'   in-place but also returns the modified object for convenience.
#'
#' @details
#' Internally, this function leverages two helper functions:
#' \itemize{
#'   \item \code{all_info_cols()}: For writing structured data tables (multiple rows and columns).
#'   \item \code{all_info_vars()}: For writing single-cell data (individual variables).
#' }
#'
#' Each of these helper functions processes metadata separately to determine:
#' \itemize{
#'   \item Which sheet the data should be written to (\code{sheet_name}).
#'   \item The starting row and column positions (\code{row_start}, \code{col_start}).
#'   \item Whether column names should be included in the Excel output (\code{col_names}, typically FALSE).
#' }
#'
#' After preparing metadata-based instructions, this function uses a custom internal utility
#' \code{pbuild()} to iteratively call \code{openxlsx2::wb_add_data()}, threading the workbook object
#' through each data insertion step. This pattern supports clean iteration to build on or
#' mutate an object over iterations with no reliance on global scope.
#'
#' This function depends on \code{pbuild()}, which is defined in the \code{listful} package.
#'
#' @examples
#' \dontrun{
#' wb <- openxlsx2::wb_load("template.xlsx")
#' metadata <- summarize_metadata("template.xlsx")
#'
#' df <- tibble::tibble(
#'   test_date = Sys.Date(),
#'   test_number = 42,
#'   projects = list(tibble::tibble(name = c("Alpha", "Beta"), value = c(100, 200)))
#' )
#'
#' wb <- write_data(wb, df, metadata$all_info[[1]])
#' openxlsx2::wb_save(wb, "populated_template.xlsx", overwrite = TRUE)
#' }
#'
#' @importFrom listful pbuild
#' @export

write_data <- function(wb, df, all_info) {

  wb_add_data_customized <- function(wb, x, sheet, start_col, start_row) {
    openxlsx2::wb_add_data(
      wb,
      sheet = sheet,
      x = x,
      start_col = start_col,
      start_row = start_row,
      col_names = FALSE,
      na.strings = ""
    )
  }

  # Write tables
  table_list <- all_info_cols(all_info, df) %>%
    select(
      x = data,
      sheet = sheet_name,
      start_col = col_start,
      start_row = row_start
    )

  wb <- wb %>%
    pbuild(table_list, wb_add_data_customized)

  # Write variables (single-cell values)
  variable_list <- all_info_vars(all_info, df) %>%
    select(
      x = data,
      sheet = sheet_name,
      start_col = col_start,
      start_row = row_start
    )

  wb <- wb %>%
    pbuild(variable_list, wb_add_data_customized)

  return(wb)
}
