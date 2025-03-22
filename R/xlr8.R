#' Populate and Style Excel Templates with Data
#'
#' Automates the workflow of reading metadata from an Excel template,
#' populating the template with user-provided data, applying original Excel formatting styles,
#' and saving the final formatted workbook. This function serves as the primary entry point
#' for creating consistently styled Excel reports from R.
#'
#' @param df Data frame containing data to populate into the Excel template. Must align with the structure defined in the Excel template metadata.
#' @param metadata_path File path to the Excel template workbook containing embedded metadata tags that specify variable, table, and column positions and formatting.
#' @param output_path File path where the final populated and styled Excel workbook will be saved.
#' @param sheets Optional character vector specifying exact sheet names to read from the template. Defaults to \code{NULL}, in which case sheets matching \code{sheets_regex} are read.
#' @param sheets_regex Regular expression pattern used to select sheets when \code{sheets} is \code{NULL}. Defaults to \code{"."}, matching all sheets.
#'
#' @return Invisibly returns \code{NULL}. The primary result is the Excel file created and saved at \code{output_path}.
#'
#' @details
#' The \code{xlr8()} function integrates multiple package components to simplify Excel-based reporting:
#'
#' \enumerate{
#'   \item **Metadata Extraction**: Reads structured metadata from an Excel template file using \code{\link{summarize_metadata}}.
#'   \item **Data Writing**: Writes provided data frame into the Excel workbook according to positions defined in the extracted metadata (\code{\link{write_data}}).
#'   \item **Style Application**: Extracts and applies original cell formatting (styles) from the Excel template onto the newly written data, preserving formatting consistency (\code{\link{apply_styles}}).
#'   \item **Output Generation**: Saves the final workbook as a formatted Excel file at the specified output location.
#' }
#'
#' This function assumes the input data (\code{df}) precisely matches the metadata-defined structure embedded within the Excel template. It provides clear error messages if misalignments occur.
#'
#' @examples
#' \dontrun{
#' df <- tibble::tibble(name = c("Alice", "Bob"), score = c(92, 85))
#'
#' xlr8(df,
#'      metadata_path = "template.xlsx",
#'      output_path = "final_report.xlsx")
#' }
#'
#' @export
xlr8 <- function (df,
                  metadata_path,
                  output_path,
                  sheets = NULL,
                  sheets_regex = "."){


  all_info <- summarize_metadata(metadata_path = metadata_path,
                                 sheets = sheets,
                                 sheets_regex = sheets_regex) %>%
    pull_cell(all_info)

  loadWorkbook(metadata_path) %>%
    write_data(df, all_info) %>%
    apply_styles(all_info,
                 df,
                 xlsx_cells(metadata_path),
                 xlsx_formats(metadata_path)
                 ) %>%
    saveWorkbook(output_path, overwrite = TRUE)

}
