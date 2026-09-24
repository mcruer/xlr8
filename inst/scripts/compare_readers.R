# =============================================================================
# compare_readers.R
#
# Point the two Excel readers at real files and report every difference.
#
#   read_excel_all()      current, tidyxl-backed
#   read_excel_all_dev()  candidate, openxlsx2-backed
#
# HOW TO RUN
#   1. Edit FILES below to point at the workbooks you want to check.
#   2. Source this whole file (in RStudio: Ctrl+Shift+S, or the Source button).
#   3. Read the report printed to the console.
#   4. Two CSVs are written next to the files, for sending on.
#
# WHAT YOU WANT TO SEE
#   real         0   the readers agree on every actual value
#   value_vs_na  0   neither reader invented or lost a value
#
#   Anything in empty_vs_na / whitespace / line_ending is expected and
#   harmless: xlr8_read() runs parse_guess_all() over everything it returns,
#   which turns "", " " and NA all into NA before tina ever sees them.
#
#   A file listed as [old_error] where the new reader succeeded is the tidyxl
#   comment bug -- a Note on a formatted-but-empty cell. That is the failure
#   this change exists to fix, so finding one here is a good result.
# =============================================================================

library(xlr8)
library(dplyr)

# ---- 1. What to compare -----------------------------------------------------
# A folder (every .xlsx / .xlsm in it):
FILES <- "C:/path/to/your/tracker/files"

# ...or an explicit set of files, if you'd rather be specific:
# FILES <- c(
#   "C:/path/to/12 - Some DSB.xlsx",
#   "C:/path/to/44 - Another DSB.xlsx"
# )

# Where to write the CSVs. Defaults to your working directory.
OUT_DIR <- getwd()

# ---- 2. Run -----------------------------------------------------------------
comparison <- compare_read_excel_all(FILES)

# ---- 3. Report --------------------------------------------------------------
print(comparison)

# ---- 4. Save ----------------------------------------------------------------
readr::write_csv(comparison$summary,
                 file.path(OUT_DIR, "reader_comparison_summary.csv"))
readr::write_csv(comparison$differences,
                 file.path(OUT_DIR, "reader_comparison_differences.csv"))

message("\nWritten to ", OUT_DIR, ":")
message("  reader_comparison_summary.csv")
message("  reader_comparison_differences.csv")

# ---- 5. Dig in --------------------------------------------------------------
# The differences that matter:
#   comparison$differences %>% filter(category %in% c("real", "value_vs_na"))
#
# Per-file detail:
#   comparison$summary %>% select(file, status, dims_match, diff_real, formulas_lost)
#
# Formulas the old reader found and the new one didn't (shared-formula cells).
# Non-empty is expected; it only matters if one of these is a tagged
# formula_location in the metadata template:
#   comparison$formulas_lost
#
# Any file where a reader failed outright:
#   comparison$summary %>% filter(status != "both_ok") %>% select(file, status, old_error, new_error)
