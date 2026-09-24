# =============================================================================
# compare_readers.R
#
# Point the two Excel readers at real files and report (1) every difference
# between them and (2) how long each one takes.
#
#   read_excel_all()      current, tidyxl-backed
#   read_excel_all_dev()  candidate, openxlsx2-backed
#
# HOW TO RUN
#   1. Check FILES below points where you want.
#   2. Source this whole file (in RStudio: Ctrl+Shift+S, or the Source button).
#   3. Read the two reports printed to the console.
#   4. Three CSVs are written to OUT_DIR, for sending on.
#
# WHAT YOU WANT TO SEE IN THE DIFF REPORT
#   real         0   the readers agree on every actual value
#   value_vs_na  0   neither reader invented or lost a value
#
#   Anything in empty_vs_na / whitespace / line_ending is expected and
#   harmless: xlr8_read() runs parse_guess_all() over everything it returns,
#   which turns "", " " and NA all into NA before anything downstream sees them.
#
#   A file listed as [old_error] where the new reader succeeded is the tidyxl
#   comment bug -- a Note on a formatted-but-empty cell. That is the failure
#   this change exists to fix, so finding one is a good result.
#
# WHAT TO EXPECT FROM THE TIMING REPORT
#   Unclear, which is why it's here. On synthetic dense sheets the new reader
#   is ~3x faster; on the TINA metadata template -- wide and sparse, the shape
#   the board files are closest to -- it is about 2x SLOWER. Real board files
#   are the only thing that settles which way it goes in production.
# =============================================================================

library(xlr8)
library(dplyr)

# ---- 1. What to compare -----------------------------------------------------
# A folder (every .xlsx / .xlsm in it):
FILES <- "B:/^TINA Package Tracker/Archive/2026-09-24 10-01-44"

# ...or an explicit set of files, if you'd rather be specific:
# FILES <- c(
#   "B:/^TINA Package Tracker/12 - Some DSB.xlsx",
#   "B:/^TINA Package Tracker/44 - Another DSB.xlsx"
# )

# Where to write the CSVs. Defaults to your working directory.
OUT_DIR <- getwd()

# How many times to read each file with each reader when timing. 3 is usually
# enough; raise it if the numbers look unstable.
REPS <- 3

# ---- 2. Do the readers agree? -----------------------------------------------
comparison <- compare_read_excel_all(FILES)
print(comparison)

# ---- 3. How fast is each one? -----------------------------------------------
# Reads every file REPS times with each reader, so this takes roughly
# 2 * REPS * (one full pass) to run. On 63 files that is not instant.
timings <- benchmark_read_excel_all(FILES, reps = REPS)
print(timings)

# ---- 4. Save ----------------------------------------------------------------
readr::write_csv(comparison$summary,
                 file.path(OUT_DIR, "reader_comparison_summary.csv"))
readr::write_csv(comparison$differences,
                 file.path(OUT_DIR, "reader_comparison_differences.csv"))
readr::write_csv(timings,
                 file.path(OUT_DIR, "reader_timings.csv"))

message("\nWritten to ", OUT_DIR, ":")
message("  reader_comparison_summary.csv")
message("  reader_comparison_differences.csv")
message("  reader_timings.csv")

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
#
# Slowest files, and which reader won each:
#   timings %>% arrange(desc(old_best)) %>% select(file, cells, old_best, new_best, ratio)
