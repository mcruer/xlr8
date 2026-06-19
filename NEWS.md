# xlr8 (development version)

## New features

* **Fan sheets.** A template sheet can now be tagged as a "fan" so that, at write
  time, it is cloned once per row of a table and each clone is renamed and
  populated with a single row of data — producing one Excel tab per record (e.g.
  one tab per project) instead of one row in a wide flat table. Reading reverses
  the process. The R-side data (`df`) is identical whether a table is rendered as
  a flat table or as fanned sheets; only the Excel tags differ.

  Tag grammar (place anywhere on the sheet; may be hidden):

  ```
  *((fan*((<table_name>*((fan_tab_name*((<naming_column>
  ```

  Fields on a fan sheet are tagged with `*((col*((<field_name>` (not `var`), each
  in its own cell, anywhere on the page. The `<naming_column>` value names each
  cloned tab and is recovered from the tab name on read (so it survives even
  though ordinary tagged cells are overwritten with data). The naming column does
  not need to appear as a visible `col` on the page, but may.

  Bidirectional: `xlr8_write()` / `xlr8_write_one()` expand fans on write, and
  `xlr8_read()` (with the new `fan_info` argument) reconstructs them on read.

  New validation rules: a fan sheet may not carry `var` tags or a normal
  table (`tbl`/`table_end`), and a given table may be fanned by at most one sheet.
  Invalid Excel tab names and tab-name collisions (including collisions with the
  still-present template sheet during expansion) error out without silent
  sanitisation.

## Known limitations / verification needed

* **Data validation (dropdowns) and conditional formatting on cloned fan sheets.**
  Whether `openxlsx2::wb_clone_worksheet()` preserves these on its own varies by
  `openxlsx2` version. `expand_fan_sheets()` captures them from the template and
  re-applies them to each clone on a best-effort basis (warning, not erroring, on
  a structural mismatch). This path reaches into `openxlsx2` internals and **must
  be confirmed empirically against the installed `openxlsx2` version** before being
  relied upon. Data validation is the higher priority of the two to preserve.

* `xlr8_read_folder()` does not yet support fan sheets, pending a `fan_info`
  column in the external `form_metadata` store.

## Bug fixes

* `xlr8_read()` now passes `keep_empty = TRUE` to `tidyr::unnest()` so that
  empty/zero-row nested tables are preserved rather than silently dropped.

* `xlr8_read()`'s internal table extraction now guards against an inverted row
  range (`row_end < row_start`), which previously could return reversed rows
  instead of an empty result.
