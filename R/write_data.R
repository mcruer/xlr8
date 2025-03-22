#' Write Data to Excel Workbook Based on Metadata Information
#'
#' Writes data from an R data frame into an Excel workbook according to
#' structured metadata. Populates the workbook with both table-based data
#' (multiple rows and columns) and single-cell variable data, based on positions
#' and structures defined in the provided metadata.
#'
#' @param wb An Excel workbook object created by \code{openxlsx::loadWorkbook()}
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
#'   \item Whether column names should be included in the Excel output (\code{colNames}, typically FALSE).
#' }
#'
#' After preparing metadata-based instructions, this function calls
#' \code{openxlsx::writeData()} iteratively (via \code{purrr::pwalk}) to populate the Excel workbook.
#'
#' @examples
#' \dontrun{
#' wb <- openxlsx::loadWorkbook("template.xlsx")
#' metadata <- summarize_metadata("template.xlsx")
#'
#' df <- tibble::tibble(name = c("A", "B"), value = c(10, 20))
#'
#' wb <- write_data(wb, df, metadata$all_info[[1]])
#' openxlsx::saveWorkbook(wb, "populated_template.xlsx", overwrite = TRUE)
#' }
#'
#' @export
write_data <- function (wb,
                        df,
                        all_info) {



  all_info_cols(all_info, df) %>%
    select (
      x = data,
      sheet = sheet_name,
      startCol = col_start,
      startRow = row_start,
      colNames
    ) %>%
    mutate(wb = list(wb)) %>%
    pwalk(openxlsx::writeData)


  all_info_vars(all_info, df) %>%
    select (
      x = data,
      sheet = sheet_name,
      startCol = col_start,
      startRow = row_start,
      colNames
    ) %>%
    mutate(wb = list(wb)) %>%
    pwalk(openxlsx::writeData)

  return(wb)

}

