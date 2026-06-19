#!/usr/bin/env Rscript
# Build a set of xlr8 *fan-sheet* templates that exercise the write -> read
# pipeline end to end. Each template here validates cleanly and round-trips.
#
# Pure openxlsx2 -- no xlr8/gplyr/listful needed just to author a template.
# Run:   Rscript dev/pipeline-test/build_templates.R [out_dir]
# or source this file and call build_all_templates("some/dir").
#
# NOTE: the ordinary flat-table path (*((tbl / *((col / *((table_end) is not
# covered here on purpose. It currently has a pre-existing validation quirk
# (a bare `*((table_end`, as documented in the vignettes and shipped in
# inst/extdata/example_metadata.xlsx, is rejected by validate_metadata as a
# "Table End Without a Matching Table"). That is unrelated to fan sheets and is
# tracked separately; see README.md in this folder.

suppressPackageStartupMessages(library(openxlsx2))

# --- authoring helpers -------------------------------------------------------
# openxlsx2's wb_add_* return the (modified) workbook; thread it through every
# call rather than relying on in-place mutation (not guaranteed across versions).

sheet <- function(wb, name) openxlsx2::wb_add_worksheet(wb, name)

put <- function(wb, sheet, text, row, col) {
  openxlsx2::wb_add_data(wb, sheet, x = text,
                         start_col = col, start_row = row, col_names = FALSE)
}

add_summary <- function(wb) {
  wb <- sheet(wb, "summary")
  put(wb, "summary", "*((var*((report_title", 1, 1)
}

add_form_sheet <- function(wb, form_id = "pipeline_v0001") {
  wb <- sheet(wb, "form")
  put(wb, "form", form_id, row = 1, col = 1)
}

# A fan sheet: the fan tag in the top-left cell (often hidden in practice),
# plus one `col` tag per field. The naming column is recovered from the tab
# name at read time, so it is NOT given a `col` tag here.
add_fan_sheet <- function(wb, template_sheet, tbl, naming_col, fields) {
  wb <- sheet(wb, template_sheet)
  wb <- put(wb, template_sheet,
            sprintf("*((fan*((%s*((fan_tab_name*((%s", tbl, naming_col), 1, 1)
  for (i in seq_along(fields)) {
    wb <- put(wb, template_sheet,
              sprintf("*((col*((%s", fields[[i]]), row = 1 + i, col = 2)
  }
  wb
}

# --- 1. fan basic: a var + one fan table -------------------------------------
build_fan_basic <- function(path) {
  wb <- openxlsx2::wb_workbook()
  wb <- add_summary(wb)
  wb <- add_fan_sheet(wb, "projects_template", "projects", "project_id",
                      c("status", "budget"))
  wb <- add_form_sheet(wb)
  openxlsx2::wb_save(wb, path, overwrite = TRUE)
  path
}

# --- 2. multi-fan: two independent fan tables --------------------------------
build_multi_fan <- function(path) {
  wb <- openxlsx2::wb_workbook()
  wb <- add_summary(wb)
  wb <- add_fan_sheet(wb, "projects_template", "projects", "project_id",
                      c("status", "budget"))
  wb <- add_fan_sheet(wb, "people_template", "people", "person_id",
                      c("role", "fte"))
  wb <- add_form_sheet(wb)
  openxlsx2::wb_save(wb, path, overwrite = TRUE)
  path
}

# --- 3. fan with a date column (exercises the date round-trip) ----------------
build_fan_dates <- function(path) {
  wb <- openxlsx2::wb_workbook()
  wb <- add_summary(wb)
  wb <- add_fan_sheet(wb, "projects_template", "projects", "project_id",
                      c("status", "start_date"))
  wb <- add_form_sheet(wb)
  openxlsx2::wb_save(wb, path, overwrite = TRUE)
  path
}

# --- 4. fan that fans out to many tabs (N-way clone + tab ordering) ----------
build_fan_many <- function(path) {
  wb <- openxlsx2::wb_workbook()
  wb <- add_summary(wb)
  wb <- add_fan_sheet(wb, "projects_template", "projects", "project_id",
                      c("status", "budget", "owner"))
  wb <- add_form_sheet(wb)
  openxlsx2::wb_save(wb, path, overwrite = TRUE)
  path
}

build_all_templates <- function(out_dir) {
  dir.create(out_dir, showWarnings = FALSE, recursive = TRUE)
  c(
    fan_basic = build_fan_basic(file.path(out_dir, "01_fan_basic.xlsx")),
    multi_fan = build_multi_fan(file.path(out_dir, "02_multi_fan.xlsx")),
    fan_dates = build_fan_dates(file.path(out_dir, "03_fan_dates.xlsx")),
    fan_many  = build_fan_many(file.path(out_dir, "04_fan_many.xlsx"))
  )
}

if (sys.nframe() == 0) {
  args <- commandArgs(trailingOnly = TRUE)
  out_dir <- if (length(args) >= 1) args[[1]] else "dev/pipeline-test/templates"
  paths <- build_all_templates(out_dir)
  cat("Wrote templates:\n"); cat(paste0("  ", paths, collapse = "\n"), "\n")
}
