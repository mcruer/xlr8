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

## Known gap — the flat-table path is intentionally not covered here

These templates only use fan sheets. The ordinary flat-table path
(`*((tbl` / `*((col` / `*((table_end`) currently has two pre-existing issues,
unrelated to fan sheets, that make it a poor fit for a clean round-trip smoke
test:

1. A **bare `*((table_end`** marker — the grammar shown in the package
   vignettes and shipped in `inst/extdata/example_metadata.xlsx` — is rejected
   by `validate_metadata()` as *"Table End Without a Matching Table"*. (You can
   reproduce this directly: `summarize_metadata("inst/extdata/example_metadata.xlsx")`
   fails validation.) The validator matches a `table_end` to a table by the
   table name extracted from the tag, but a bare `*((table_end` extracts no
   name, so it never matches.
2. A **named `*((table_end*((<tbl>`** marker validates, but the marker row is
   read back as a trailing garbage data row unless the written data happens to
   fill exactly up to it (the read range is `row_start:row_end` inclusive, with
   no trailing-NA / marker-row trimming).

Both are tracked as separate cleanup items for the flat-table read path and are
out of scope for the fan-sheet feature.
