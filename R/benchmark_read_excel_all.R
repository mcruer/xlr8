#' Benchmark read_excel_all() against read_excel_all_dev()
#'
#' Times both readers over the same workbooks, repeating each read so the
#' numbers are stable enough to compare. Separate from
#' [compare_read_excel_all()], whose per-file timings are single-run and only
#' meant as a rough indication.
#'
#' @param paths Character vector of workbook paths, or a single directory, in
#'   which case every \code{.xlsx} / \code{.xlsm} file in it is used (files
#'   whose names contain \code{~} are skipped, as Excel lock files do).
#' @param reps Number of times to read each file with each reader. Defaults to
#'   3. Raise it for small files, where a single read is short enough that
#'   timer granularity shows.
#' @param sheets,sheets_regex Passed through to both readers unchanged.
#' @param quiet If \code{FALSE} (the default) prints progress as it goes.
#'
#' @return An object of class \code{xlr8_reader_benchmark}: a tibble with one
#'   row per file giving cell count, and the fastest and median elapsed seconds
#'   for each reader, plus the speed ratio. A reader that errors on a file
#'   gives \code{NA} timings for that file rather than ending the run.
#'
#' @details
#' Reported times are wall clock from \code{Sys.time()}. The fastest of
#' \code{reps} runs is the more reliable figure -- it is the one least polluted
#' by garbage collection and other background work -- and the median is given
#' alongside it as a sanity check. \code{ratio} is old/new, so above 1 means
#' the new reader is faster.
#'
#' Expect the gap to be modest on small workbooks and to widen with size. Most
#' of what [read_excel_all()] does is the reshaping after the read -- the
#' \code{expand_grid()} and \code{pivot_wider()} -- which is identical in both
#' readers, so the parse is only part of the total.
#'
#' @seealso [compare_read_excel_all()], which checks the two readers agree.
#'
#' @examples
#' \dontrun{
#' bm <- benchmark_read_excel_all("C:/path/to/tracker/files")
#' bm
#' }
#'
#' @export
benchmark_read_excel_all <- function(paths,
                                     reps = 3,
                                     sheets = NULL,
                                     sheets_regex = ".",
                                     quiet = FALSE) {

  if (length(paths) == 1 && dir.exists(paths)) {
    paths <- list.files(paths, pattern = "\\.xls[xm]$", full.names = TRUE)
    paths <- paths[!stringr::str_detect(basename(paths), "~")]
  }
  if (length(paths) == 0) stop("No files to benchmark.")

  out <- purrr::map_dfr(seq_along(paths), function(i) {
    if (!quiet) message("[", i, "/", length(paths), "] ", basename(paths[i]))
    benchmark_one_file_dev(paths[i], reps, sheets, sheets_regex)
  })

  structure(out, class = c("xlr8_reader_benchmark", class(out)))
}


#' Time both readers on a single file
#'
#' @param path Path to one workbook.
#' @param reps Number of reads per reader.
#' @param sheets,sheets_regex Passed to both readers.
#'
#' @return A one-row tibble of timings for this file.
#'
#' @keywords internal
benchmark_one_file_dev <- function(path, reps, sheets, sheets_regex) {

  time_reader <- function(f) {
    secs <- rep(NA_real_, reps)
    cells <- NA_integer_
    for (i in seq_len(reps)) {
      # Collect before timing, not during: without this a large file leaves
      # enough garbage that the *next* file's timings absorb the collection,
      # which is easily big enough to reverse the comparison.
      gc(verbose = FALSE)
      t0 <- Sys.time()
      res <- tryCatch(f(path, sheets = sheets, sheets_regex = sheets_regex),
                      error = function(e) NULL)
      if (is.null(res)) return(list(secs = NA_real_, med = NA_real_,
                                    cells = NA_integer_))
      secs[i] <- as.numeric(difftime(Sys.time(), t0, units = "secs"))
      if (i == 1) cells <- as.integer(nrow(res) * (ncol(res) - 2))
    }
    list(secs = min(secs), med = stats::median(secs), cells = cells)
  }

  old <- time_reader(read_excel_all_tidyxl)
  new <- time_reader(read_excel_all_dev)

  tibble::tibble(
    file       = basename(path),
    cells      = dplyr::coalesce(old$cells, new$cells),
    reps       = reps,
    old_best   = round(old$secs, 4),
    old_median = round(old$med, 4),
    new_best   = round(new$secs, 4),
    new_median = round(new$med, 4),
    ratio      = round(old$secs / new$secs, 2)
  )
}


#' @export
print.xlr8_reader_benchmark <- function(x, ...) {
  cat("read_excel_all() vs read_excel_all_dev() - timings\n")
  cat(strrep("-", 56), "\n")
  cat("files:", nrow(x), " | reps per reader:", x$reps[1], "\n")
  cat("(best of", x$reps[1], "runs; ratio is old/new, >1 means new is faster)\n\n")

  print(tibble::as_tibble(x) %>%
          dplyr::select(file, cells, old_best, new_best, ratio),
        n = nrow(x))

  ok <- dplyr::filter(x, !is.na(old_best), !is.na(new_best))
  if (nrow(ok) > 0) {
    to <- sum(ok$old_best); tn <- sum(ok$new_best)
    cat("\ntotal over", nrow(ok), "files both readers could read:\n")
    cat("   read_excel_all()     ", sprintf("%7.3f s", to), "\n")
    cat("   read_excel_all_dev() ", sprintf("%7.3f s", tn), "\n")
    cat("   ratio                ", sprintf("%7.2f", to / tn),
        if (to > tn) " (new is faster)\n" else " (old is faster)\n")
    if (sum(ok$cells, na.rm = TRUE) > 0) {
      cat("\nthroughput (cells/sec):  old",
          format(round(sum(ok$cells) / to), big.mark = ","),
          " | new", format(round(sum(ok$cells) / tn), big.mark = ","), "\n")
    }
  }

  failed <- dplyr::filter(x, is.na(old_best) | is.na(new_best))
  if (nrow(failed) > 0) {
    cat("\nnot timed (a reader errored -- see compare_read_excel_all() for why):\n")
    for (i in seq_len(nrow(failed))) {
      cat("  ", failed$file[i],
          if (is.na(failed$old_best[i])) "[read_excel_all() failed]"
          else "[read_excel_all_dev() failed]", "\n")
    }
  }
  invisible(x)
}
