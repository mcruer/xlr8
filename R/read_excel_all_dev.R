#' Read All Cells from Excel Sheets into a Structured Tibble (openxlsx2 backend)
#'
#' A drop-in alternative to [read_excel_all()] that reads the workbook with
#' \code{openxlsx2} instead of \code{tidyxl}. Same arguments, same return shape,
#' same cloaked attributes. Intended to be compared against [read_excel_all()]
#' on real files (see [compare_read_excel_all()]) before replacing it.
#'
#' @param path File path to the Excel workbook.
#' @param sheets Optional character vector specifying exact sheet names to read.
#'   Defaults to \code{NULL}, in which case sheets are matched using
#'   \code{sheets_regex}.
#' @param sheets_regex Regular expression pattern to select sheets when
#'   \code{sheets} is \code{NULL}. Defaults to \code{"."}, matching all sheets.
#'
#' @return Identical in shape to [read_excel_all()]: a wide tibble with
#'   \code{sheet_name}, \code{row}, and \code{x1}, \code{x2}, ... columns,
#'   carrying cloaked \code{formulas} and \code{wb} attributes.
#'
#' @details
#' Why this exists: \code{tidyxl::xlsx_cells(include_blank_cells = FALSE)}
#' overruns its output vectors when a cell comment sits on a cell that has a
#' \code{<c>} element but no value (a formatted-but-empty cell), failing with
#' \code{attempt to set index N/N in SET_STRING_ELT}. Any analyst who adds a
#' Note to an empty cell breaks the read. \code{openxlsx2} does not have this
#' bug, is already a dependency of this package, and is already loaded by
#' [read_excel_all()] for the cloaked workbook -- so this backend also removes
#' a redundant second parse of every file.
#'
#' Three differences between the two libraries are normalised here so the
#' output matches [read_excel_all()] cell for cell:
#'
#' \itemize{
#'   \item \code{openxlsx2} returns the raw XML text of a formula
#'     (\code{&gt;}); \code{tidyxl} unescapes entities (\code{>}).
#'   \item A formula cell whose result is the empty string reads as \code{""}
#'     in \code{tidyxl} but \code{NA} in \code{openxlsx2}.
#'   \item \code{tidyxl} preserves CRLF inside cells; \code{openxlsx2}
#'     normalises to LF.
#' }
#'
#' \strong{Known behavioural difference.} \code{tidyxl} propagates shared
#' formulas: where Excel stores a formula once against a master cell and lets
#' neighbouring cells inherit it, \code{tidyxl} reconstructs the inherited text
#' and shifts its relative references. \code{openxlsx2} returns \code{NA} for
#' those inheriting cells. On the TINA production metadata template this costs
#' 550 formulas, none of which sits at a tagged \code{formula_location}, so
#' \code{summarize_metadata()} is unaffected -- but see
#' [check_formula_locations()], which turns that from a silent failure into an
#' error.
#'
#' @seealso [read_excel_all()], the tidyxl-backed original;
#'   [compare_read_excel_all()], which diffs the two over real files.
#'
#' @examples
#' \dontrun{
#' read_excel_all_dev("workbook.xlsx")
#' read_excel_all_dev("workbook.xlsx", sheets = c("Data", "Summary"))
#' }
#'
#' @export
read_excel_all_dev <- function(path, sheets = NULL, sheets_regex = ".") {

  wb <- openxlsx2::wb_load(path)
  all_names <- unname(wb$get_sheet_names())

  if (is.null(sheets)) {
    sheet_names <- gplyr::str_filter(all_names, sheets_regex)
    if (length(sheet_names) == 0) {
      stop("The sheets_regex argument didn't match any sheets in the workbook.")
    }
  } else {
    problems <- setdiff(sheets, all_names)
    if (length(problems) > 0) {
      stop(stringr::str_c("The following sheets are not in the workbook: ",
                          stringr::str_c(problems, sep = ", ", collapse = ", ")))
    }
    sheet_names <- sheets
  }

  initial <- purrr::map_dfr(sheet_names, read_one_sheet_dev, wb = wb,
                            all_names = all_names)

  out <- initial %>%
    dplyr::select(-formula) %>%
    dplyr::group_by(sheet_name) %>%
    gplyr::quicks(c(row, col), max) %>%
    dplyr::mutate(df = purrr::map2(row, col,
                                   ~ tidyr::expand_grid(row = 1:.x, col = 1:.y))) %>%
    dplyr::select(-row, -col) %>%
    tidyr::unnest(df) %>%
    dplyr::full_join(initial %>% dplyr::select(-formula)) %>%
    tidyr::pivot_wider(names_from = col, values_from = cell_contents) %>%
    dplyr::rename_with(~ stringr::str_c("x", .x), -c(sheet_name, row)) %>%
    dplyr::mutate(sheet_name = as.character(sheet_name)) %>%
    dplyr::arrange(sheet_name) %>%
    suppressMessages()

  gplyr::cloak(out, list(formulas = initial, wb = wb))
}


#' Read one sheet into long form for read_excel_all_dev()
#'
#' @param sheet_name Name of the sheet to read.
#' @param wb An \code{openxlsx2} workbook object from \code{wb_load()}.
#' @param all_names Character vector of every sheet name in \code{wb}, in order.
#'
#' @return A tibble of \code{sheet_name}, \code{row}, \code{col},
#'   \code{cell_contents}, \code{formula}, one row per non-empty cell.
#'
#' @keywords internal
read_one_sheet_dev <- function(sheet_name, wb, all_names) {

  cc <- wb$worksheets[[which(all_names == sheet_name)]]$sheet_data$cc
  if (is.null(cc) || nrow(cc) == 0) return(tibble::tibble())

  rows <- as.integer(cc$row_r)
  cols <- openxlsx2::col2int(cc$c_r)

  # Anchor the read at A1 so leading blank rows and columns are preserved --
  # without explicit dims, wb_to_df() starts at the first non-empty cell, which
  # would shift every position the metadata relies on.
  dims <- openxlsx2::wb_dims(rows = seq_len(max(rows)), cols = seq_len(max(cols)))
  d <- openxlsx2::wb_to_df(wb, sheet = sheet_name, col_names = FALSE, dims = dims,
                           skip_empty_rows = FALSE, skip_empty_cols = FALSE)

  values <- tibble::tibble(
    row = rep(as.integer(rownames(d)), times = ncol(d)),
    col = rep(openxlsx2::col2int(colnames(d)), each = nrow(d)),
    # tidyxl keeps CRLF inside cells; openxlsx2 normalises to LF.
    cell_contents = stringr::str_replace_all(
      as.character(unlist(d, use.names = FALSE)), "(?<!\r)\n", "\r\n")
  )

  formulas <- tibble::tibble(
    row = rows,
    col = cols,
    formula = unescape_xml_dev(dplyr::na_if(cc$f, ""))
  )

  values %>%
    dplyr::left_join(formulas, by = c("row", "col")) %>%
    dplyr::filter(!is.na(cell_contents) | !is.na(formula)) %>%
    # A formula cell whose result is the empty string reads as "" in tidyxl but
    # NA in openxlsx2. Match tidyxl.
    dplyr::mutate(cell_contents = dplyr::if_else(
      is.na(cell_contents) & !is.na(formula), "", cell_contents)) %>%
    dplyr::mutate(sheet_name = sheet_name, .before = 1)
}


#' Unescape XML entities in formula text
#'
#' \code{openxlsx2} hands back the raw XML text of a formula, in which
#' \code{<}, \code{>}, \code{&}, \code{"} and \code{'} are escaped.
#' \code{tidyxl} unescapes them. Match tidyxl.
#'
#' @param x Character vector of formula text.
#'
#' @return \code{x} with XML entities replaced by the characters they stand for.
#'
#' @keywords internal
unescape_xml_dev <- function(x) {
  x <- stringr::str_replace_all(x, "&lt;", "<")
  x <- stringr::str_replace_all(x, "&gt;", ">")
  x <- stringr::str_replace_all(x, "&quot;", "\"")
  x <- stringr::str_replace_all(x, "&apos;", "'")
  # &amp; last, so an escaped entity such as &amp;lt; survives as &lt;
  stringr::str_replace_all(x, "&amp;", "&")
}
