#' Retrieve File Information for a Single Path
#'
#' This function returns a tibble with details about files in a given folder path.
#'
#' @param folder_path Character. The path of the folder from which files' information is to be extracted.
#' @param file_type Character. Type or extension of the files to be considered. Default is "." (all files).
#' @param recursive Logical. Should files from subdirectories also be included? Default is FALSE.
#' @param filter_out_tilda Logical. Should files with names containing a tilde (~) be excluded? Default is TRUE.
#'
#' @return A tibble containing columns: file (name of the file), path (complete path of the file),
#' size (file size), isdir (is it a directory?), and mtime (last modification time).
#'
#' @importFrom fs path
#' @importFrom stringr str_detect str_c str_sub
#' @importFrom purrr map
#' @importFrom tibble tibble
#' @importFrom dplyr select filter
#' @importFrom tidyr unnest
#' @importFrom gplyr filter_in_str
#' @examples
#' \dontrun{
#' file_tibble_single_path(folder_path = "./my_directory")
#' }
#' @keywords internal
file_tibble_single_path <- function (folder_path,
                                     file_type = ".",
                                     recursive = FALSE,
                                     filter_out_tilda = TRUE) {

  if (str_sub(folder_path, start = -1L) != "/") {
    folder_path <- str_c(folder_path, "/")
  }

  tibble(
    file = list.files(folder_path, recursive = recursive),
    path = path(folder_path, file),
    info = map(path, ~ .x |>
                 file.info() |>
                 tibble() |>
                 select(size, isdir, mtime))
  ) |>
    unnest(info) |>
    filter_in_str(file, str_c(file_type, "$")) |>
    filter(!(filter_out_tilda & str_detect(file, "~")))
}


#' Re-Nest Metadata
#'
#' This function aggregates file information from multiple folder paths into a single tibble.
#'
#' @param df A dataframe containing columns to be contained within the metadata column.
#'
#' @return A tibble with a nested metadata column.
#'
#' @importFrom tidyr nest
#' @importFrom dplyr any_of
#' @export
nest_metadata <- function (df) {
  metadata_columns <- c(
    "folder_paths",
    "file_type",
    "recursive",
    "filter_out_tilda",
    "size",
    "isdir",
    "mtime",
    "update_needed_raw",
    "update_needed_form",
    "form_name_version",
    "sheets",
    "sheets_regex"
  )

  nest(df, metadata = any_of(metadata_columns))
}

#' Retrieve File Information for Multiple Paths
#'
#' This function aggregates file information from multiple folder paths into a single tibble.
#'
#' @param folder_paths Character vector. Paths of the folders from which files' information is to be extracted.
#' @param file_type Character. Type or extension of the files to be considered. Default is "." (all files).
#' @param recursive Logical. Should files from subdirectories also be included? Default is FALSE.
#' @param filter_out_tilda Logical. Should files with names containing a tilde (~) be excluded? Default is TRUE.
#'
#' @return A tibble with one row per unique file-path combination, with columns: folder_paths (list of all folder paths provided),
#' file_type (type of the file considered), recursive (were subdirectories included?), filter_out_tilda (were files with tildes excluded?),
#' file (name of the file), path (complete path of the file), and metadata (a list column containing details like size, isdir, and mtime).
#'
#' @importFrom purrr map
#' @importFrom dplyr bind_rows mutate
#' @export
file_tibble <- function (folder_paths,
                         file_type = ".",
                         recursive = FALSE,
                         filter_out_tilda = TRUE) {

  bind_rows(
    map(folder_paths,
        file_tibble_single_path,
        file_type = file_type,
        recursive = recursive,
        filter_out_tilda = filter_out_tilda)
  ) |>
    mutate(
      folder_paths = list(folder_paths),
      file_type = file_type,
      recursive = recursive,
      filter_out_tilda = filter_out_tilda,
      update_needed_raw = TRUE,
      update_needed_form = TRUE,
      .before = 1
    ) |>
    nest_metadata()
}





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
