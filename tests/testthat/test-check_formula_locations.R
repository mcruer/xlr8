# Tests for check_formula_locations(), the guard summarize_metadata() uses to
# catch a tagged formula_location that yielded no formula.

tagged <- function(...) {
  tibble::tribble(
    ~sheet_name, ~formula_location, ~formula,
    ...
  )
}

test_that("it passes silently when every tagged location has a formula", {
  ok <- tagged(
    "Monthly Update", "G11",  "IF(A10=\"\", \"\", 1)",
    "Monthly Update", "BX11", "SUM(AO10:AR10)"
  )
  expect_no_error(check_formula_locations(ok))
  expect_silent(check_formula_locations(ok))
})

test_that("it returns its input invisibly, so it can sit in a pipe", {
  ok <- tagged("Sheet1", "A1", "SUM(B1:B2)")
  expect_identical(withVisible(check_formula_locations(ok))$visible, FALSE)
  expect_identical(check_formula_locations(ok), ok)
})

test_that("it stops and names every cell that came back without a formula", {
  bad <- tagged(
    "Monthly Update", "G11",  "IF(A10=\"\", \"\", 1)",
    "Monthly Update", "BX11", NA_character_,
    "Errors",         "CF11", NA_character_
  )

  expect_error(check_formula_locations(bad), "Monthly Update!BX11", fixed = TRUE)
  expect_error(check_formula_locations(bad), "Errors!CF11", fixed = TRUE)
  expect_error(check_formula_locations(bad), "2 tagged formula_location", fixed = TRUE)
  # The message has to say what to do about it, since the failure is a silent
  # data error rather than a crash.
  expect_error(check_formula_locations(bad), "shared-formula master")
})

test_that("an empty set of tagged locations is not an error", {
  expect_no_error(check_formula_locations(tagged()))
})
