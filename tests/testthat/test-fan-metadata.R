# Tests for the "fan sheet" feature that do not require a live Excel workbook.
# Round-trip tests that need openxlsx2 file I/O live in a separate file and are
# skipped when openxlsx2 is unavailable.

test_that("target_extraction distinguishes fan from fan_tab_name (prefix collision)", {
  tag <- "*((fan*((projects*((fan_tab_name*((project_id"
  expect_equal(xlr8:::target_extraction(tag, "fan"), "projects")
  expect_equal(xlr8:::target_extraction(tag, "fan_tab_name"), "project_id")
})

test_that("validate_excel_sheet_names flags illegal names and passes valid ones", {
  ok <- xlr8:::validate_excel_sheet_names(c("Project A", "P-101", "2025_Q1"))
  expect_equal(nrow(ok), 0)

  bad <- xlr8:::validate_excel_sheet_names(c(
    "fine",
    paste(rep("x", 32), collapse = ""), # 32 chars > 31
    "bad/name",
    "bad:name",
    "bad[name]",
    "",
    NA_character_
  ))
  # every entry except "fine" should be reported
  expect_setequal(
    bad$tab_name,
    c(paste(rep("x", 32), collapse = ""), "bad/name", "bad:name", "bad[name]",
      "", NA_character_)
  )
})

test_that("collapse_fan_sheets rebuilds one row per fan-instance sheet", {
  raw_df <- tibble::tribble(
    ~sheet_name, ~row, ~x1, ~x2,
    "P1", 1L, "*((fan*((projects*((fan_tab_name*((project_id", NA_character_,
    "P1", 2L, "status:", "open",
    "P2", 1L, "*((fan*((projects*((fan_tab_name*((project_id", NA_character_,
    "P2", 2L, "status:", "closed"
  )

  all_info <- tibble::tibble(
    row_id = "form projects status",
    sheet_name = "template",
    tbl = "projects",
    row_start = 2, row_end = 2,
    col_start = 2, col_end = 2,
    col_name = "status",
    formula_location = NA_character_
  )

  fan_info <- tibble::tibble(
    sheet_name = "template", row = 1, col = 1,
    tbl = "projects", fan_tab_name_col = "project_id"
  )

  res <- xlr8:::collapse_fan_sheets(raw_df, all_info, fan_info)
  expect_true("projects" %in% names(res))

  proj <- res$projects[[1]]
  expect_equal(nrow(proj), 2)
  expect_true(all(c("status", "project_id") %in% names(proj)))
  # naming column is recovered from the tab (sheet) name, not a cell
  expect_setequal(proj$project_id, c("P1", "P2"))
  expect_equal(proj$status[proj$project_id == "P1"], "open")
  expect_equal(proj$status[proj$project_id == "P2"], "closed")
})

test_that("collapse_fan_sheets yields a zero-row tibble when no instances exist", {
  # A workbook with no surviving fan tags (e.g. zero projects at write time).
  raw_df <- tibble::tribble(
    ~sheet_name, ~row, ~x1,
    "Summary", 1L, "nothing here"
  )

  all_info <- tibble::tibble(
    row_id = "form projects status",
    sheet_name = "template",
    tbl = "projects",
    row_start = 2, row_end = 2,
    col_start = 2, col_end = 2,
    col_name = "status",
    formula_location = NA_character_
  )

  fan_info <- tibble::tibble(
    sheet_name = "template", row = 1, col = 1,
    tbl = "projects", fan_tab_name_col = "project_id"
  )

  res <- xlr8:::collapse_fan_sheets(raw_df, all_info, fan_info)
  proj <- res$projects[[1]]
  expect_equal(nrow(proj), 0)
  expect_true(all(c("status", "project_id") %in% names(proj)))
})
