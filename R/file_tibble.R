
#' Read and Extract Structured Data from Excel Files in a Folder
#'
#' Scans a folder for `.xlsx` files, reads them using `read_excel_all()`,
#' extracts the form version from each file, retrieves corresponding metadata
#' from the `form_metadata` database table, and optionally extracts structured data
#' using `xlr8_read()`.
#'
#' @param path Character. Folder path to search for Excel files.
#' @param recursive Logical. Whether to search subdirectories recursively. Default is `FALSE`.
#' @param filter_out_tilda Logical. Whether to exclude files with a tilde (`~`) in the name. Default is `TRUE`.
#' @param sheets Optional character vector specifying exact sheet names to read. Passed to `read_excel_all()`.
#' @param sheet_regex Regular expression to match sheet names if `sheets` is `NULL`. Default is `"."`.
#' @param extract Logical. Whether to extract structured data using `xlr8_read()`. If `FALSE`, only metadata is attached.
#'
#' @return A tibble where each row represents a processed Excel file. Columns include:
#' \describe{
#'   \item{file}{File name.}
#'   \item{path}{Full file path.}
#'   \item{raw_df}{Raw data read from the Excel file (nested).}
#'   \item{form_name}{Form name extracted from the hidden "form" sheet.}
#'   \item{form_metadata}{Metadata tibble for the form (nested), if found in the database.}
#'   \item{...}{If `extract = TRUE`, additional columns for each variable/table defined in the form metadata. Tables are nested tibbles.}
#' }
#'
#' @details
#' - Excel files must include a hidden sheet named `"form"` with a version string (e.g., `"my_form_v0001"`) in cell A1.
#' - Metadata must be stored in a database table named `"form_metadata"` (populated via `summarize_metadata()`).
#' - Progress bars are shown if `.progress` is used (requires `purrr >= 1.0.0`).
#'
#' @importFrom purrr map map_chr map2
#' @importFrom dplyr mutate select left_join distinct
#' @importFrom tidyr unnest
#' @importFrom tibble tibble
#' @importFrom databased file_tibble
#' @export
xlr8_read_folder <- function (path,
                              recursive = FALSE,
                              filter_out_tilda = TRUE,
                              sheets = NULL,
                              sheet_regex = ".",
                              extract = TRUE) {
  out <- file_tibble(
    path,
    file_type = "xlsx",
    recursive = recursive,
    filter_out_tilda = filter_out_tilda
  ) %>%
    mutate(
      raw_df = map(path, read_excel_all, sheets, sheet_regex),
      form_name = map_chr(raw_df, extract_form_name, .progress = "Reading in Raw Files")
    )

  if (extract) {
    #An interim step is needed here so we don't query the database unnecessarily.
    metadata_forms <- out %>%
      select(form_name) %>%
      unique() %>%
      mutate(form_metadata = map(
        form_name,
        ~ ezql_table("form_metadata") %>% filter(form == .x)
      ))

    out <- out %>%
      left_join(metadata_forms) %>%
      mutate(data = map2(raw_df, form_metadata, xlr8_read, .progress = "Extracting Data")) %>%
      unnest(data)
  }
  return(out)

  #Just as an FYI these are the columns nested within the metadata.

  # metadata_columns <- c(
  #   "folder_paths",
  #   "file_type",
  #   "recursive",
  #   "filter_out_tilda",
  #   "size",
  #   "isdir",
  #   "mtime",
  #   "update_needed_raw",
  #   "update_needed_form",
  #   "form_name_version",
  #   "sheets",
  #   "sheets_regex"
  # )

}
