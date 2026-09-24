# Tests for the comparison and benchmark harnesses.

fixture <- function(name) testthat::test_path("fixtures", name)

test_that("differences are classified into the right categories", {
  classify <- classify_diff_dev

  # Agreement is not a difference.
  expect_true(is.na(classify("a", "a")))
  expect_true(is.na(classify(NA_character_, NA_character_)))

  # The two that matter.
  expect_equal(classify("1899-12-31", "1899-12-30"), "real")
  expect_equal(classify("a value", NA_character_), "value_vs_na")
  expect_equal(classify(NA_character_, "a value"), "value_vs_na")

  # The three that don't.
  expect_equal(classify("", NA_character_), "empty_vs_na")
  expect_equal(classify(NA_character_, ""), "empty_vs_na")
  expect_equal(classify(" ", NA_character_), "empty_vs_na")
  expect_equal(classify("a\r\nb", "a\nb"), "line_ending")
  expect_equal(classify(" a ", "a"), "whitespace")

  # Vectorised.
  expect_equal(classify(c("a", "", "x"), c("a", NA, "y")),
               c(NA, "empty_vs_na", "real"))
})

test_that("a reader erroring is recorded, not thrown", {
  # tidyxl cannot read this one, which is the whole point of the change.
  cmp <- compare_read_excel_all(fixture("comment-on-blank-cell.xlsx"), quiet = TRUE)

  expect_s3_class(cmp, "xlr8_reader_comparison")
  expect_equal(cmp$summary$status, "old_error")
  expect_match(cmp$summary$old_error, "SET_STRING_ELT")
  expect_true(is.na(cmp$summary$new_error))
})

test_that("comparing a clean file finds nothing that matters", {
  cmp <- compare_read_excel_all(fixture("value-types.xlsx"), quiet = TRUE)

  expect_equal(cmp$summary$status, "both_ok")
  expect_equal(cmp$summary$diff_real, 0L)
  expect_equal(cmp$summary$diff_value_vs_na, 0L)
  expect_true(cmp$summary$dims_match)
})

test_that("the serial-zero disagreement is reported as a real difference", {
  # This is the one thing the two backends genuinely disagree about, so the
  # harness must surface it rather than absorb it into a harmless category.
  cmp <- compare_read_excel_all(fixture("date-serial-edges.xlsx"), quiet = TRUE)

  expect_gt(cmp$summary$diff_real, 0L)
  expect_true(any(cmp$differences$category == "real"))
  expect_equal(cmp$summary$diff_value_vs_na, 0L)
})

test_that("lost shared formulas are reported", {
  cmp <- compare_read_excel_all(fixture("shared-formula.xlsx"), quiet = TRUE)

  expect_equal(cmp$summary$formulas_lost, 2L)   # B2 and B3 inherit B1's formula
  expect_equal(nrow(cmp$formulas_lost), 2L)
  expect_true(all(cmp$formulas_lost$col == 2))
})

test_that("a directory compares every workbook in it", {
  cmp <- compare_read_excel_all(fixture(""), quiet = TRUE)

  expect_gt(nrow(cmp$summary), 1)
  expect_true(all(c("value-types.xlsx", "offset-origin.xlsx") %in% cmp$summary$file))
  # Every file either compares cleanly or is one tidyxl can't read.
  expect_true(all(cmp$summary$status %in% c("both_ok", "old_error")))
  ok <- cmp$summary[cmp$summary$status == "both_ok", ]
  expect_equal(sum(ok$diff_value_vs_na), 0L)
})

test_that("printing a comparison does not error", {
  cmp <- compare_read_excel_all(fixture("value-types.xlsx"), quiet = TRUE)
  expect_output(print(cmp), "read_excel_all")
})

test_that("an empty set of paths is an error rather than an empty report", {
  expect_error(compare_read_excel_all(character()), "No files to compare")
})


test_that("the benchmark times both readers and reports the ratio", {
  bm <- benchmark_read_excel_all(fixture("value-types.xlsx"), reps = 1, quiet = TRUE)

  expect_s3_class(bm, "xlr8_reader_benchmark")
  expect_equal(nrow(bm), 1L)
  expect_gt(bm$cells, 0)
  expect_gt(bm$old_best, 0)
  expect_gt(bm$new_best, 0)
  expect_equal(bm$ratio, round(bm$old_best / bm$new_best, 2))
})

test_that("the benchmark records NA rather than failing on an unreadable file", {
  bm <- benchmark_read_excel_all(fixture("comment-on-blank-cell.xlsx"),
                                 reps = 1, quiet = TRUE)
  expect_true(is.na(bm$old_best))
  expect_gt(bm$new_best, 0)
  expect_true(is.na(bm$ratio))
  expect_output(print(bm), "not timed")
})
