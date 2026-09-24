# Tests for read_excel_all() and its openxlsx2 backend.
#
# These assert absolute expected values rather than "matches tidyxl", so they
# keep their meaning once read_excel_all_tidyxl() and the tidyxl dependency are
# removed.

fixture <- function(name) testthat::test_path("fixtures", name)

# Pull one cell out of the wide grid: sheet + Excel row + Excel column index.
cell_at <- function(df, sheet, row, col) {
  df[[paste0("x", col)]][df$sheet_name == sheet & df$row == row]
}


# --- the bug this backend exists to fix -------------------------------------

test_that("a comment on a formatted-but-empty cell does not break the read", {
  path <- fixture("comment-on-blank-cell.xlsx")

  expect_no_error(read_excel_all(path))

  out <- read_excel_all(path)
  expect_s3_class(out, "data.frame")

  # All of the real data is there.
  expect_equal(cell_at(out, "Data", 2, 1), "alpha")
  expect_equal(cell_at(out, "Data", 2, 2), "1.5")
  expect_equal(cell_at(out, "Data", 2, 3), "2026-03-25")
  expect_equal(cell_at(out, "Data", 3, 4), "FALSE")

  # The commented cell is F4 -- column 6, holding nothing. A cell with neither
  # value nor formula contributes nothing, so the grid stops at column 4 rather
  # than stretching out to an empty column 6.
  expect_equal(names(out), c("sheet_name", "row", "x1", "x2", "x3", "x4"))
  expect_true(all(is.na(unlist(out[out$row == 4, -(1:2)]))))
})

test_that("the superseded tidyxl backend still fails on that workbook", {
  skip_if_not_installed("tidyxl")
  # Not a demand that tidyxl stay broken -- it documents why read_excel_all()
  # moved off it. If tidyxl is ever fixed this test should be deleted, not
  # worked around.
  expect_error(
    read_excel_all_tidyxl(fixture("comment-on-blank-cell.xlsx")),
    "SET_STRING_ELT"
  )
})


# --- positional fidelity ----------------------------------------------------

test_that("leading blank rows and columns are preserved", {
  out <- read_excel_all(fixture("offset-origin.xlsx"))

  # Data starts at D5. The grid must still start at row 1, column 1.
  expect_equal(min(out$row), 1)
  expect_true("x1" %in% names(out))
  expect_equal(sort(out$row), 1:7)

  # D5 must land at row 5, column 4 -- not row 1, column 1.
  expect_equal(cell_at(out, "Offset", 5, 4), "hdrD")
  expect_equal(cell_at(out, "Offset", 5, 6), "10")
  expect_equal(cell_at(out, "Offset", 7, 5), "e7")

  # ...and the empty region really is empty.
  expect_true(all(is.na(out$x1)))
  expect_true(all(is.na(cell_at(out, "Offset", 1, 4))))
})


# --- value rendering --------------------------------------------------------

test_that("cell values survive conversion to character without loss", {
  out <- read_excel_all(fixture("value-types.xlsx"))
  v <- function(row) cell_at(out, "Data", row, 1)

  expect_equal(v(1),  "00123")                 # text, leading zeros kept
  expect_equal(v(2),  "0.30000000000000004")   # full float precision
  expect_equal(v(3),  "1234567890123456")      # large integer, not 1.23e+15
  expect_equal(v(4),  "0.0000123")             # not re-formatted
  expect_equal(v(5),  "text with  spaces")     # internal spacing kept
  expect_equal(v(6),  "2026-03-25 12:15:00")   # datetime
  expect_equal(v(7),  "#DIV/0!")               # error value
  expect_equal(v(8),  "TRUE")                  # boolean
  expect_equal(v(9),  "1e-9")                  # scientific notation verbatim
  expect_equal(v(10), "0.125")                 # percent-formatted stays raw
  expect_equal(v(11), "100000000")
  expect_equal(v(12), "2026-01-01")            # a date held as text
})


# --- formulas ---------------------------------------------------------------

test_that("formula text is XML-unescaped", {
  formulas <- gplyr::uncloak(read_excel_all(fixture("shared-formula-empty-result.xlsx")))$formulas
  master <- formulas$formula[!is.na(formulas$formula)][1]

  expect_equal(master, 'IF(A1>0,"","bad")')
  expect_false(grepl("&gt;|&lt;|&amp;", master))
})

test_that("a formula whose result is the empty string reads as \"\", not NA", {
  out <- read_excel_all(fixture("shared-formula-empty-result.xlsx"))
  # B1 is the shared-formula master, so its formula is visible and its empty
  # result is rendered as "".
  expect_identical(cell_at(out, "E", 1, 2), "")
})

test_that("shared formulas are not propagated to the cells that inherit them", {
  # Characterisation, not endorsement. tidyxl reconstructs inherited formulas
  # and shifts their references; openxlsx2 does not, so cells inheriting a
  # shared formula come back NA. summarize_metadata() guards against a tagged
  # formula_location landing on one -- see check_formula_locations().
  formulas <- gplyr::uncloak(read_excel_all(fixture("shared-formula.xlsx")))$formulas
  b <- formulas[formulas$col == 2, ]
  b <- b[order(b$row), ]

  expect_equal(b$formula[1], "A1*2")          # master carries the text
  expect_true(all(is.na(b$formula[-1])))      # inheriting cells do not

  # If this ever starts failing because openxlsx2 gained propagation, that is
  # good news: relax the guard, don't relax the test.
})


# --- dates ------------------------------------------------------------------

test_that("ordinary dates and datetimes are exact", {
  out <- read_excel_all(fixture("date-serial-edges.xlsx"))
  expect_equal(cell_at(out, "T", 8, 1), "2023-03-15 12:00:00")
})

test_that("serial zero renders as 1899-12-30, not a real date", {
  # Excel serial 0 is not a date -- Excel shows it as 1900-01-00. It is the only
  # value the tidyxl and openxlsx2 backends ever disagreed about across 63
  # production workbooks (tidyxl said 1899-12-31). Pinned so the choice is
  # deliberate rather than incidental.
  out <- read_excel_all(fixture("date-serial-edges.xlsx"))
  expect_equal(cell_at(out, "T", 5, 1), "1899-12-30")
})

test_that("time-formatted cells render as times", {
  out <- read_excel_all(fixture("date-serial-edges.xlsx"))
  expect_equal(cell_at(out, "T", 2, 1), "12:00:00")
  expect_equal(cell_at(out, "T", 3, 1), "06:00:00")
})


# --- sheet selection --------------------------------------------------------

test_that("sheets and sheets_regex select sheets, and bad input errors", {
  path <- fixture("value-types.xlsx")

  expect_equal(unique(read_excel_all(path, sheets = "Data")$sheet_name), "Data")
  expect_equal(unique(read_excel_all(path, sheets_regex = "^Dat")$sheet_name), "Data")

  expect_error(read_excel_all(path, sheets = "Nope"), "not in the workbook")
  expect_error(read_excel_all(path, sheets_regex = "zzz"), "didn't match any sheets")
})


# --- the wrapper ------------------------------------------------------------

test_that("read_excel_all() is read_excel_all_dev()", {
  strip <- function(x) {
    x <- as.data.frame(x)
    for (a in setdiff(names(attributes(x)), c("names", "row.names", "class"))) {
      attr(x, a) <- NULL
    }
    x
  }
  for (f in c("value-types.xlsx", "offset-origin.xlsx", "shared-formula.xlsx")) {
    expect_identical(strip(read_excel_all(fixture(f))),
                     strip(read_excel_all_dev(fixture(f))),
                     info = f)
  }
})

test_that("the result carries the cloaked formulas and workbook", {
  cloaked <- gplyr::uncloak(read_excel_all(fixture("shared-formula.xlsx")))

  expect_named(cloaked, c("formulas", "wb"), ignore.order = TRUE)
  expect_true(all(c("sheet_name", "row", "col", "cell_contents", "formula") %in%
                    names(cloaked$formulas)))
  expect_s3_class(cloaked$wb, "wbWorkbook")
})
