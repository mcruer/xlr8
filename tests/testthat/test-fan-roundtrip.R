# End-to-end round trip for fan sheets. Requires openxlsx2 file I/O (and, in
# particular, exercises wb_clone_worksheet / wb_remove_worksheet, whose fidelity
# for data validation / conditional formatting must be confirmed separately).

build_fan_template <- function() {
  path <- tempfile(fileext = ".xlsx")
  wb <- openxlsx2::wb_workbook()

  # A normal var on a non-fan sheet, to exercise the mixed case.
  wb <- openxlsx2::wb_add_worksheet(wb, "summary")
  wb <- openxlsx2::wb_add_data(wb, "summary", x = "*((var*((report_title",
                               start_col = 1, start_row = 1, col_names = FALSE)

  # The fan sheet: a fan tag plus two scattered col tags.
  wb <- openxlsx2::wb_add_worksheet(wb, "projects_template")
  wb <- openxlsx2::wb_add_data(
    wb, "projects_template",
    x = "*((fan*((projects*((fan_tab_name*((project_id",
    start_col = 1, start_row = 1, col_names = FALSE)
  wb <- openxlsx2::wb_add_data(wb, "projects_template", x = "*((col*((status",
                               start_col = 2, start_row = 2, col_names = FALSE)
  wb <- openxlsx2::wb_add_data(wb, "projects_template", x = "*((col*((budget",
                               start_col = 2, start_row = 3, col_names = FALSE)

  # The hidden form sheet.
  wb <- openxlsx2::wb_add_worksheet(wb, "form")
  wb <- openxlsx2::wb_add_data(wb, "form", x = "testform_v0001",
                               start_col = 1, start_row = 1, col_names = FALSE)

  openxlsx2::wb_save(wb, path, overwrite = TRUE)
  path
}

test_that("fan sheets round trip to the same nested-tibble shape", {
  skip_if_not_installed("openxlsx2")
  skip_if_not_installed("tidyxl")

  template_path <- build_fan_template()
  out_path <- tempfile(fileext = ".xlsx")

  df <- tibble::tibble(
    report_title = "Q1 Report",
    projects = list(tibble::tibble(
      project_id = c("Alpha", "Beta"),
      status     = c("open", "closed"),
      budget     = c(100, 200)
    ))
  )

  xlr8_write_one(df, output_path = out_path,
                 metadata_path = template_path, overwrite = TRUE)

  # The template sheet is gone; one sheet per project exists.
  sheet_names <- as.character(openxlsx2::wb_get_sheet_names(
    openxlsx2::wb_load(out_path)))
  expect_false("projects_template" %in% sheet_names)
  expect_true(all(c("Alpha", "Beta") %in% sheet_names))

  meta <- summarize_metadata(template_path)
  raw <- read_excel_all(out_path)
  res <- xlr8_read(raw,
                   all_info = meta$all_info[[1]],
                   fan_info = meta$fan_info[[1]])

  expect_equal(res$report_title, "Q1 Report")

  proj <- res$projects[[1]]
  expect_equal(nrow(proj), 2)
  expect_setequal(proj$project_id, c("Alpha", "Beta"))
  expect_equal(proj$status[proj$project_id == "Alpha"], "open")
  expect_equal(proj$budget[proj$project_id == "Beta"], 200)
})

test_that("zero-row fan table warns and drops the template sheet", {
  skip_if_not_installed("openxlsx2")
  skip_if_not_installed("tidyxl")

  template_path <- build_fan_template()
  out_path <- tempfile(fileext = ".xlsx")

  df <- tibble::tibble(
    report_title = "Empty Report",
    projects = list(tibble::tibble(
      project_id = character(),
      status     = character(),
      budget     = numeric()
    ))
  )

  expect_warning(
    xlr8_write_one(df, output_path = out_path,
                   metadata_path = template_path, overwrite = TRUE),
    "zero rows"
  )

  sheet_names <- as.character(openxlsx2::wb_get_sheet_names(
    openxlsx2::wb_load(out_path)))
  expect_false("projects_template" %in% sheet_names)

  meta <- summarize_metadata(template_path)
  raw <- read_excel_all(out_path)
  res <- xlr8_read(raw,
                   all_info = meta$all_info[[1]],
                   fan_info = meta$fan_info[[1]])
  expect_equal(nrow(res$projects[[1]]), 0)
})

test_that("colliding fan tab names error before mutating the workbook", {
  skip_if_not_installed("openxlsx2")
  skip_if_not_installed("tidyxl")

  template_path <- build_fan_template()
  out_path <- tempfile(fileext = ".xlsx")

  df <- tibble::tibble(
    report_title = "Dup Report",
    projects = list(tibble::tibble(
      project_id = c("Same", "Same"),
      status     = c("open", "closed"),
      budget     = c(1, 2)
    ))
  )

  expect_error(
    xlr8_write_one(df, output_path = out_path,
                   metadata_path = template_path, overwrite = TRUE),
    "duplicate sheet names"
  )
})
