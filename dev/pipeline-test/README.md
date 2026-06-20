# Fan-sheet pipeline test harness

A small, self-contained exercise of the **fan-sheet** write → read pipeline,
end to end, across several templates. It complements the unit tests in
`tests/testthat/` (which cover the same feature in finer-grained pieces) with a
"does the whole thing actually round-trip from a real `.xlsx`" smoke test you
can run by hand and eyeball.

## Files

- `build_templates.R` — authors the test templates with `openxlsx2` only
  (no `xlr8`/`gplyr`/`listful` needed just to build a template).
- `run_pipeline.R` — for each scenario: writes a filled workbook from a data
  frame, reads it back, and asserts the round-tripped data matches.
- `templates/` — the generated `.xlsx` templates, committed so you can open
  them in Excel and see the tags. They are regenerated into a temp dir by
  `run_pipeline.R`, so they are never strictly required to run the test.

## Running

From the package root:

```sh
Rscript dev/pipeline-test/run_pipeline.R
```

It loads the working-tree source with `pkgload::load_all()` (falling back to
`library(xlr8)`), so it always tests the code you have checked out. To just
(re)generate the templates:

```sh
Rscript dev/pipeline-test/build_templates.R dev/pipeline-test/templates
```

## Scenarios covered

| Template | What it exercises |
|----------|-------------------|
| `01_fan_basic` | One fan table + a plain `var`. The core promise: a fanned table round-trips to the **same nested tibble** a flat table would produce. |
| `02_multi_fan` | Two independent fan tables in one workbook (different naming columns, different field sets). |
| `03_fan_dates` | A date column survives write → read as a `Date`. |
| `04_fan_many`  | A fan that expands to many tabs (N-way clone), all rows/columns recovered. |
| (reuses `01`) | Zero-row fan: warns, drops the template sheet, reads back an empty table. |
| (reuses `01`) | Colliding tab names: errors **before** mutating the workbook. |
| `05_flat_basic` | An ordinary flat table (`*((tbl` / `*((col` / a bare `*((table_end`) round-trips, with no trailing marker-row leak. |

## Fixed: two pre-existing flat-table bugs

While building this harness, two pre-existing issues turned up in the ordinary
flat-table path (`*((tbl` / `*((col` / `*((table_end`), unrelated to fan
sheets. Both are now fixed (see `R/summarize_metadata.R`):

1. A **bare `*((table_end`** marker — the grammar shown in the package
   vignettes and shipped in `inst/extdata/example_metadata.xlsx` — used to be
   rejected by `validate_metadata()` as *"Table End Without a Matching
   Table"*, because the validator matched a `table_end` to a table by the name
   extracted from the tag, and a bare `*((table_end` extracts no name. It's
   now treated as ending every table on its own sheet, rather than requiring a
   name match.
2. A **named `*((table_end*((<tbl>`** marker used to read back as a trailing
   garbage data row, because the resolved `row_end` pointed at the marker's
   own row rather than the last real data row above it. `row_end` is now
   stepped back one row so the marker row is excluded from the read range.

`05_flat_basic` exercises both: it uses the bare-marker grammar, and asserts no
marker row leaks into the round-tripped table.
