#!/usr/bin/env Rscript
# End-to-end exercise of the xlr8 fan-sheet write -> read pipeline across a set
# of templates. For each scenario it: writes a filled workbook from a data
# frame, reads it back, and checks the round-tripped data matches expectations
# (including the nested-tibble shape a flat table would have produced).
#
# Usage (from the package root):
#   Rscript dev/pipeline-test/run_pipeline.R
#
# It loads the package with pkgload::load_all() so it always runs against the
# working-tree source. Templates are (re)generated into a temp dir from
# build_templates.R, so the committed .xlsx files are never required -- though
# they are committed alongside this script so you can open them in Excel.

suppressWarnings(suppressMessages({
  if (requireNamespace("pkgload", quietly = TRUE)) {
    pkgload::load_all(quiet = TRUE)
  } else {
    library(xlr8)
  }
  library(tibble)
}))

here <- tryCatch(dirname(sys.frame(1)$ofile), error = function(e) NA)
if (is.na(here) || !nzchar(here)) here <- "dev/pipeline-test"
source(file.path(here, "build_templates.R"))

# --- tiny check harness ------------------------------------------------------
.failures <- 0L
ok <- function(cond, msg) {
  status <- if (isTRUE(cond)) "PASS" else { .failures <<- .failures + 1L; "FAIL" }
  cat(sprintf("  [%s] %s\n", status, msg))
}
eq <- function(a, b, msg) ok(isTRUE(all.equal(a, b)), msg)

# Canonicalise a nested table for order-independent comparison: coerce to a
# plain data.frame (the fan read returns a tibble; an input may be either) and
# sort both rows (by key) and columns (by name) so neither ordering matters.
canon <- function(df, key) {
  df <- as.data.frame(df, stringsAsFactors = FALSE)
  df <- df[order(df[[key]]), order(names(df)), drop = FALSE]
  rownames(df) <- NULL
  df
}

write_then_read <- function(tpl, df) {
  out <- tempfile(fileext = ".xlsx")
  meta <- summarize_metadata(tpl, quiet = TRUE)
  xlr8_write_one(df, output_path = out, metadata_path = tpl, overwrite = TRUE)
  raw <- read_excel_all(out)
  res <- xlr8_read(raw, all_info = meta$all_info[[1]], fan_info = meta$fan_info[[1]])
  list(out = out, res = res,
       sheets = as.character(openxlsx2::wb_get_sheet_names(openxlsx2::wb_load(out))))
}

tpl_dir <- tempfile("xlr8_pipeline_templates")
paths <- build_all_templates(tpl_dir)
cat("Templates generated in:", tpl_dir, "\n\n")

# ============================================================================
cat("1. fan_basic: one fan table round-trips to a nested tibble\n")
# ============================================================================
{
  df <- tibble(
    report_title = "Q1 Report",
    projects = list(tibble(
      project_id = c("Alpha", "Beta"),
      status     = c("open", "closed"),
      budget     = c(100, 200)
    ))
  )
  r <- write_then_read(paths[["fan_basic"]], df)

  ok(!("projects_template" %in% r$sheets), "template sheet removed")
  ok(all(c("Alpha", "Beta") %in% r$sheets), "one tab per project created")
  eq(r$res$report_title, "Q1 Report", "var round-trips")

  got <- canon(r$res$projects[[1]], "project_id")
  want <- canon(as.data.frame(df$projects[[1]]), "project_id")
  eq(got, want, "projects nested tibble equals the input (shape + values + types)")
}

# ============================================================================
cat("\n2. multi_fan: two independent fan tables in one workbook\n")
# ============================================================================
{
  df <- tibble(
    report_title = "Multi",
    projects = list(tibble(
      project_id = c("Alpha", "Beta"),
      status     = c("open", "closed"),
      budget     = c(100, 200)
    )),
    people = list(tibble(
      person_id = c("Pat", "Sam", "Lee"),
      role      = c("PM", "Eng", "Design"),
      fte       = c(1, 0.5, 0.8)
    ))
  )
  r <- write_then_read(paths[["multi_fan"]], df)

  ok(all(c("Alpha", "Beta", "Pat", "Sam", "Lee") %in% r$sheets),
     "tabs created for both fans")
  ok(!any(c("projects_template", "people_template") %in% r$sheets),
     "both template sheets removed")
  eq(canon(r$res$projects[[1]], "project_id"),
     canon(as.data.frame(df$projects[[1]]), "project_id"),
     "projects table round-trips")
  eq(canon(r$res$people[[1]], "person_id"),
     canon(as.data.frame(df$people[[1]]), "person_id"),
     "people table round-trips")
}

# ============================================================================
cat("\n3. fan_dates: a date column survives the round-trip as a Date\n")
# ============================================================================
{
  df <- tibble(
    report_title = "Dates",
    projects = list(tibble(
      project_id = c("Alpha", "Beta"),
      status     = c("open", "closed"),
      start_date = as.Date(c("2025-01-15", "2025-03-02"))
    ))
  )
  r <- write_then_read(paths[["fan_dates"]], df)
  got <- canon(r$res$projects[[1]], "project_id")
  ok(inherits(got$start_date, "Date"), "start_date reads back as Date")
  eq(got$start_date, sort(df$projects[[1]]$start_date), "date values preserved")
}

# ============================================================================
cat("\n4. fan_many: fans out to many tabs, preserves all rows/cols\n")
# ============================================================================
{
  ids <- sprintf("P%02d", 1:6)
  df <- tibble(
    report_title = "Many",
    projects = list(tibble(
      project_id = ids,
      status     = rep(c("open", "closed"), length.out = 6),
      budget     = (1:6) * 100,
      owner      = letters[1:6]
    ))
  )
  r <- write_then_read(paths[["fan_many"]], df)
  ok(all(ids %in% r$sheets), "a tab exists for every record")
  got <- canon(r$res$projects[[1]], "project_id")
  eq(nrow(got), 6, "all 6 records recovered")
  eq(got, canon(as.data.frame(df$projects[[1]]), "project_id"),
     "every column round-trips")
}

# ============================================================================
cat("\n5. zero-row fan: warns, drops template, reads an empty table\n")
# ============================================================================
{
  df <- tibble(
    report_title = "Empty",
    projects = list(tibble(
      project_id = character(), status = character(), budget = numeric()
    ))
  )
  got_warning <- FALSE
  out <- tempfile(fileext = ".xlsx")
  withCallingHandlers(
    xlr8_write_one(df, output_path = out,
                   metadata_path = paths[["fan_basic"]], overwrite = TRUE),
    warning = function(w) {
      if (grepl("zero rows", conditionMessage(w))) got_warning <<- TRUE
      invokeRestart("muffleWarning")
    }
  )
  ok(got_warning, "zero-row fan emits a warning")
  sheets <- as.character(openxlsx2::wb_get_sheet_names(openxlsx2::wb_load(out)))
  ok(!("projects_template" %in% sheets), "template sheet still removed")

  meta <- summarize_metadata(paths[["fan_basic"]], quiet = TRUE)
  res <- xlr8_read(read_excel_all(out),
                   all_info = meta$all_info[[1]], fan_info = meta$fan_info[[1]])
  eq(nrow(res$projects[[1]]), 0, "empty fan reads back as a zero-row tibble")
}

# ============================================================================
cat("\n6. colliding tab names: errors before mutating the workbook\n")
# ============================================================================
{
  df <- tibble(
    report_title = "Dup",
    projects = list(tibble(
      project_id = c("Same", "Same"),
      status     = c("open", "closed"),
      budget     = c(1, 2)
    ))
  )
  out <- tempfile(fileext = ".xlsx")
  err <- tryCatch({
    xlr8_write_one(df, output_path = out,
                   metadata_path = paths[["fan_basic"]], overwrite = TRUE)
    NULL
  }, error = function(e) conditionMessage(e))
  ok(!is.null(err) && grepl("duplicate sheet names", err),
     "duplicate fan tab names raise a clear error")
  ok(!file.exists(out), "no output file written on the failed run")
}

# ============================================================================
cat("\n7. flat_basic: an ordinary flat table round-trips (bare table_end)\n")
# ============================================================================
{
  df <- tibble(
    report_title = "Flat",
    projects = list(tibble(
      project_id = c("Alpha", "Beta"),
      status     = c("open", "closed"),
      budget     = c(100, 200)
    ))
  )
  r <- write_then_read(paths[["flat_basic"]], df)
  eq(r$res$report_title, "Flat", "var round-trips")
  got <- canon(r$res$projects[[1]], "project_id")
  want <- canon(as.data.frame(df$projects[[1]]), "project_id")
  eq(nrow(got), 2, "no trailing table_end marker row leaks into the data")
  eq(got, want, "flat table round-trips (shape + values + types)")
}

cat("\n----------------------------------------------------------------\n")
if (.failures == 0L) {
  cat("All pipeline checks passed.\n")
} else {
  cat(sprintf("%d pipeline check(s) FAILED.\n", .failures))
  quit(status = 1)
}
