#' Compare read_excel_all() against read_excel_all_dev() over real files
#'
#' Runs both readers over the same workbooks and reports every difference, so
#' the openxlsx2 backend can be validated against the tidyxl one on real data
#' before it replaces it.
#'
#' @param paths Character vector of workbook paths, or a single directory, in
#'   which case every \code{.xlsx} / \code{.xlsm} file in it is used (files
#'   whose names contain \code{~} are skipped, as Excel lock files do).
#' @param sheets,sheets_regex Passed through to both readers unchanged.
#' @param max_examples Maximum number of differing cells to keep per file.
#'   Counts in the summary are always complete; this only caps the detail table
#'   so one badly-mismatched file can't produce a million rows. Defaults to 50.
#' @param quiet If \code{FALSE} (the default) prints progress as it goes.
#'
#' @return An object of class \code{xlr8_reader_comparison}: a list of three
#'   tibbles.
#'   \describe{
#'     \item{\code{summary}}{One row per file: whether each reader succeeded,
#'       grid dimensions, cells compared, differences found by category,
#'       formula counts, and elapsed seconds for each reader.}
#'     \item{\code{differences}}{Up to \code{max_examples} differing cells per
#'       file, with the value each reader gave and a category.}
#'     \item{\code{formulas_lost}}{Formulas the tidyxl reader found that the
#'       openxlsx2 reader did not -- expected for cells that inherit a shared
#'       formula. Empty is ideal; non-empty is not necessarily a problem unless
#'       one of these cells is a tagged \code{formula_location}.}
#'   }
#'
#' @details
#' Differences are classified so that the signal isn't buried in noise:
#' \describe{
#'   \item{\code{real}}{The two readers disagree about an actual value. This is
#'     the category that matters -- anything here needs explaining.}
#'   \item{\code{value_vs_na}}{One reader found a value, the other found
#'     nothing. Also matters.}
#'   \item{\code{empty_vs_na}}{One gave \code{""}, the other \code{NA}. Harmless
#'     in practice: \code{xlr8_read()} runs \code{parse_guess_all()}, which
#'     turns both into \code{NA}.}
#'   \item{\code{whitespace}}{Equal once leading/trailing whitespace is trimmed.
#'     Also collapsed to \code{NA} by \code{parse_guess_all()}.}
#'   \item{\code{line_ending}}{Equal once CRLF and LF are normalised.}
#' }
#'
#' Either reader erroring is recorded rather than thrown, so one bad file
#' doesn't end the run. A file where \code{status} is \code{old_error} and the
#' new reader succeeded is the tidyxl comment bug -- the thing this change is
#' meant to fix.
#'
#' @seealso [read_excel_all_dev()]
#'
#' @examples
#' \dontrun{
#' cmp <- compare_read_excel_all("C:/path/to/tracker/files")
#' cmp                                    # printed report
#' cmp$summary                            # per-file table
#' dplyr::filter(cmp$differences, category == "real")
#' }
#'
#' @export
compare_read_excel_all <- function(paths,
                                   sheets = NULL,
                                   sheets_regex = ".",
                                   max_examples = 50,
                                   quiet = FALSE) {

  if (length(paths) == 1 && dir.exists(paths)) {
    paths <- list.files(paths, pattern = "\\.xls[xm]$", full.names = TRUE)
    paths <- paths[!stringr::str_detect(basename(paths), "~")]
  }
  if (length(paths) == 0) stop("No files to compare.")

  results <- purrr::map(seq_along(paths), function(i) {
    if (!quiet) message("[", i, "/", length(paths), "] ", basename(paths[i]))
    compare_one_file_dev(paths[i], sheets, sheets_regex, max_examples)
  })

  out <- list(
    summary       = purrr::map_dfr(results, "summary"),
    differences   = purrr::map_dfr(results, "differences"),
    formulas_lost = purrr::map_dfr(results, "formulas_lost")
  )
  structure(out, class = c("xlr8_reader_comparison", "list"))
}


#' Compare the two readers on a single file
#'
#' @param path Path to one workbook.
#' @param sheets,sheets_regex Passed to both readers.
#' @param max_examples Cap on differing cells retained for this file.
#'
#' @return A list of \code{summary}, \code{differences} and
#'   \code{formulas_lost} tibbles for this one file.
#'
#' @keywords internal
compare_one_file_dev <- function(path, sheets, sheets_regex, max_examples) {

  file <- basename(path)
  run <- function(f) {
    t0 <- Sys.time()
    value <- tryCatch(f(path, sheets = sheets, sheets_regex = sheets_regex),
                      error = function(e) structure(conditionMessage(e),
                                                    class = "reader_error"))
    list(value = value,
         secs  = as.numeric(difftime(Sys.time(), t0, units = "secs")),
         ok    = !inherits(value, "reader_error"))
  }

  old <- run(read_excel_all_tidyxl)
  new <- run(read_excel_all_dev)

  status <- dplyr::case_when(
    old$ok  &  new$ok ~ "both_ok",
    !old$ok &  new$ok ~ "old_error",
    old$ok  & !new$ok ~ "new_error",
    TRUE              ~ "both_error"
  )

  empty <- tibble::tibble()
  base <- tibble::tibble(
    file = file, status = status,
    old_error = if (old$ok) NA_character_ else as.character(old$value),
    new_error = if (new$ok) NA_character_ else as.character(new$value),
    secs_old = round(old$secs, 3), secs_new = round(new$secs, 3)
  )

  if (status != "both_ok") {
    return(list(summary = base, differences = empty, formulas_lost = empty))
  }

  cells <- dplyr::full_join(
    reader_long_dev(old$value), reader_long_dev(new$value),
    by = c("sheet_name", "row", "col"), suffix = c("_old", "_new")
  ) %>%
    dplyr::mutate(category = classify_diff_dev(value_old, value_new)) %>%
    dplyr::filter(!is.na(category))

  f_old <- gplyr::uncloak(old$value)$formulas
  f_new <- gplyr::uncloak(new$value)$formulas
  has_formula <- function(d) dplyr::filter(d, !is.na(formula))
  lost <- dplyr::anti_join(has_formula(f_old), has_formula(f_new),
                           by = c("sheet_name", "row", "col"))

  counts <- table(factor(cells$category,
                         levels = c("real", "value_vs_na", "empty_vs_na",
                                    "whitespace", "line_ending")))

  summary <- base %>% dplyr::mutate(
    sheets_old      = dplyr::n_distinct(f_old$sheet_name),
    sheets_new      = dplyr::n_distinct(f_new$sheet_name),
    dims_old        = paste(dim(old$value), collapse = " x "),
    dims_new        = paste(dim(new$value), collapse = " x "),
    dims_match      = identical(dim(old$value), dim(new$value)),
    cells_compared  = nrow(reader_long_dev(old$value)),
    diff_real       = as.integer(counts[["real"]]),
    diff_value_vs_na = as.integer(counts[["value_vs_na"]]),
    diff_empty_vs_na = as.integer(counts[["empty_vs_na"]]),
    diff_whitespace = as.integer(counts[["whitespace"]]),
    diff_line_ending = as.integer(counts[["line_ending"]]),
    formulas_old    = sum(!is.na(f_old$formula)),
    formulas_new    = sum(!is.na(f_new$formula)),
    formulas_lost   = nrow(lost)
  )

  differences <- cells %>%
    dplyr::arrange(match(category, c("real", "value_vs_na", "empty_vs_na",
                                     "whitespace", "line_ending"))) %>%
    utils::head(max_examples) %>%
    dplyr::transmute(file, sheet_name, row, col, category,
                     old = substr(value_old, 1, 200),
                     new = substr(value_new, 1, 200))

  list(summary = summary,
       differences = differences,
       formulas_lost = lost %>% dplyr::mutate(file = file, .before = 1))
}


#' Pivot a reader's wide output to one row per cell
#'
#' @param df A tibble as returned by [read_excel_all()] or
#'   [read_excel_all_dev()].
#'
#' @return A tibble of \code{sheet_name}, \code{row}, \code{col}, \code{value}.
#'
#' @keywords internal
reader_long_dev <- function(df) {
  df %>%
    tibble::as_tibble() %>%
    dplyr::mutate(dplyr::across(-c(sheet_name, row), as.character)) %>%
    tidyr::pivot_longer(-c(sheet_name, row), names_to = "col",
                        values_to = "value") %>%
    dplyr::mutate(col = as.integer(stringr::str_remove(col, "^x")))
}


#' Classify a pair of cell values as a kind of difference, or no difference
#'
#' @param old,new Character vectors of the same length, one from each reader.
#'
#' @return A character vector: \code{NA} where the two agree, otherwise the
#'   category of disagreement.
#'
#' @keywords internal
classify_diff_dev <- function(old, new) {
  nl <- function(x) stringr::str_replace_all(x, "\r\n", "\n")
  both_na  <- is.na(old) & is.na(new)
  one_na   <- xor(is.na(old), is.na(new))
  present  <- !is.na(old) & !is.na(new)
  blank_vs_na <- one_na &
    ((is.na(old) & !is.na(new) & trimws(new) == "") |
       (is.na(new) & !is.na(old) & trimws(old) == ""))

  dplyr::case_when(
    both_na                                          ~ NA_character_,
    present & old == new                             ~ NA_character_,
    blank_vs_na                                      ~ "empty_vs_na",
    one_na                                           ~ "value_vs_na",
    present & nl(old) == nl(new)                     ~ "line_ending",
    present & trimws(old) == trimws(new)             ~ "whitespace",
    TRUE                                             ~ "real"
  )
}


#' @export
print.xlr8_reader_comparison <- function(x, ...) {
  s <- x$summary
  cat("read_excel_all() vs read_excel_all_dev()\n")
  cat(strrep("-", 56), "\n")
  cat("files compared:", nrow(s), "\n\n")

  cat("status:\n")
  for (st in names(table(s$status))) {
    cat("  ", format(st, width = 12), table(s$status)[[st]], "\n")
  }

  ok <- dplyr::filter(s, status == "both_ok")
  if (nrow(ok) > 0) {
    cat("\ncells compared:", format(sum(ok$cells_compared), big.mark = ","), "\n")
    cat("differences:\n")
    cat("   real            ", sum(ok$diff_real), "   <- must be 0\n")
    cat("   value_vs_na     ", sum(ok$diff_value_vs_na), "   <- must be 0\n")
    cat("   empty_vs_na     ", sum(ok$diff_empty_vs_na), "   (harmless)\n")
    cat("   whitespace      ", sum(ok$diff_whitespace), "   (harmless)\n")
    cat("   line_ending     ", sum(ok$diff_line_ending), "   (harmless)\n")
    cat("\ngrid dimensions match in", sum(ok$dims_match), "of", nrow(ok), "files\n")
    cat("formulas: old", sum(ok$formulas_old), "| new", sum(ok$formulas_new),
        "| lost", sum(ok$formulas_lost), "(shared-formula cells; see $formulas_lost)\n")
    cat("\nelapsed: old", round(sum(ok$secs_old), 1), "s | new",
        round(sum(ok$secs_new), 1), "s\n")
  }

  bad <- dplyr::filter(s, status != "both_ok")
  if (nrow(bad) > 0) {
    cat("\nfiles where a reader failed:\n")
    for (i in seq_len(nrow(bad))) {
      cat("  ", bad$file[i], "[", bad$status[i], "]\n")
      msg <- dplyr::coalesce(bad$old_error[i], bad$new_error[i])
      cat("      ", substr(msg, 1, 120), "\n")
    }
  }

  hard <- dplyr::filter(x$differences, category %in% c("real", "value_vs_na"))
  if (nrow(hard) > 0) {
    cat("\nfirst differences needing explanation:\n")
    print(utils::head(hard, 10))
  }
  invisible(x)
}
