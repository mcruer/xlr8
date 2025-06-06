validate_data_to_write <- function(df, all_info) {

  all_info <- all_info %>%
    filter_in_na(formula_location)

  make_check_key <- function(level1, level2) {
    paste0(level1, tidyr::replace_na(level2, ""))
  }

  flatten_nested_tibble <- function(df) {
    purrr::imap_dfr(df, function(col, name) {
      if (!is.list(col)) {
        tibble::tibble(level1 = name, level2 = NA_character_)
      } else {
        purrr::map_dfr(col, function(x) {
          if (tibble::is_tibble(x)) {
            tibble::tibble(level1 = name, level2 = names(x))
          } else {
            tibble::tibble(level1 = name, level2 = NA_character_)
          }
        })
      }
    })
  }

  required <- all_info %>%
    dplyr::mutate(check = make_check_key(tbl, col_name))

  available <- flatten_nested_tibble(df) %>%
    dplyr::mutate(check = make_check_key(level1, level2)) %>%
    dplyr::distinct(check)

  missing_data_check <- dplyr::anti_join(required, available, by = "check") %>%
    dplyr::select(sheet_name, tbl, col_name, row_start, col_start)

  if (nrow(missing_data_check) > 0) {
    print(missing_data_check)
    stop("See tibble above. Data provided is not available for these required output variables.")
  }

  invisible(NULL)
}


#' Write One Excel Report from Data and Metadata
#'
#' Writes a single Excel workbook using a provided template, data frame, and metadata.
#' The function populates data, applies formatting styles, and saves the result.
#' Supports reuse of pre-parsed metadata and preloaded workbooks to enable efficient batching.
#'
#' @param df A tibble of data to be written. Tables should be list-columns.
#' @param output_path File path where the formatted Excel workbook will be saved.
#' @param metadata_path (Optional) Path to the Excel template file with embedded metadata tags.
#' @param all_info (Optional) Pre-parsed metadata tibble from `summarize_metadata()` (specifically the `all_info` component).
#' @param wb (Optional) A workbook object loaded using `wb_load()`. Required if `metadata_path` is not used.
#' @param sheets Optional character vector of sheet names to read when parsing metadata.
#' @param sheets_regex Optional regular expression for matching sheets. Default is `"."`.
#' @param overwrite Logical. Whether to overwrite existing files. Default is `FALSE`.
#'
#' @return Invisibly returns `NULL`. The primary result is the saved Excel workbook.
#'
#' @examples
#' \dontrun{
#' # Using only the template path
#' xlr8_write_one(df, output_path = "out.xlsx", metadata_path = "template.xlsx")
#'
#' # Reusing parsed metadata and workbook
#' all_info <- summarize_metadata("template.xlsx") %>% pull_cell(all_info)
#' wb <- wb_load("template.xlsx")
#' xlr8_write_one(df, output_path = "out.xlsx", all_info = all_info, wb = wb)
#' }
#'
#' @importFrom dplyr pull
#' @importFrom gplyr pull_cell
#' @importFrom openxlsx2 wb_load wb_save
#' @export
xlr8_write_one <- function(df,
                           output_path,
                           metadata_path = NULL,
                           all_info = NULL,
                           wb = NULL,
                           sheets = NULL,
                           sheets_regex = ".",
                           overwrite = FALSE) {

  stopifnot(is.data.frame(df), is.character(output_path))

  if (is.null(all_info)) {
    if (is.null(metadata_path)) {
      stop("Must provide either `all_info` or `metadata_path`.")
    }
    all_info <- summarize_metadata(
      metadata_path = metadata_path,
      sheets = sheets,
      sheets_regex = sheets_regex
    ) %>% pull_cell(all_info)
  }

  if (is.null(wb)) {
    if (is.null(metadata_path)) {
      stop("Must provide either `wb` or `metadata_path`.")
    }
    wb <- wb_load(metadata_path)
  }


  #Check if the df contains all the necessary data to write to the sheet.
  validate_data_to_write(df, all_info)

  wb <- wb %>%
    write_data(df, all_info) %>%
    apply_styles(df, all_info) %>%
    add_formulas(df, all_info)

  wb_save(
    wb,
    file = output_path,
    overwrite = overwrite
  )

  invisible(NULL)
}


#' Write Multiple Excel Reports from a List-Tibble
#'
#' Writes multiple Excel workbooks by mapping over a tibble where each row contains
#' the data and output path. This function wraps `xlr8_write_one()` and is designed
#' for batch generation of styled Excel files.
#'
#' @param df A tibble with at least two columns:
#'   - `df`: a list-column of data frames to be written.
#'   - `output_path`: character column with file paths.
#' @param metadata_path (Optional) Path to the Excel template to use across all rows.
#'   Ignored if `all_info` or `wb` is supplied.
#' @param all_info (Optional) Pre-parsed metadata from `summarize_metadata()`.
#' @param wb (Optional) A preloaded workbook object from `wb_load()`.
#' @param sheets Optional character vector of sheet names to read when parsing metadata.
#' @param sheets_regex Regular expression to match sheet names if `sheets` is NULL.
#' @param overwrite Logical. Whether to overwrite existing files. Default is `FALSE`.
#'
#' @return Invisibly returns the input tibble `df`.
#'
#' @importFrom dplyr pull
#' @importFrom purrr pwalk
#' @importFrom gplyr pull_cell
#' @export
xlr8_write <- function(df,
                       metadata_path = NULL,
                       all_info = NULL,
                       wb = NULL,
                       sheets = NULL,
                       sheets_regex = ".",
                       overwrite = FALSE) {

  stopifnot("df" %in% names(df), "output_path" %in% names(df))

  # Load metadata if needed
  if (is.null(all_info)) {
    if (is.null(metadata_path)) stop("Must supply either `all_info` or `metadata_path`.")
    all_info <- summarize_metadata(metadata_path, sheets, sheets_regex) %>%
      pull_cell(all_info)
  }

  # Load workbook if needed
  if (is.null(wb)) {
    if (is.null(metadata_path)) stop("Must supply either `wb` or `metadata_path`.")
    wb <- wb_load(metadata_path)
  }

  pwalk(df, function(df, output_path) {
    xlr8_write_one(
      df = df,
      output_path = output_path,
      all_info = all_info,
      wb = wb,
      overwrite = overwrite
    )
  })

  invisible(df)
}
