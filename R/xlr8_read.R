#' Read Structured Data from Tagged Excel Forms
#'
#' Extracts structured data from Excel workbooks that have been pre-tagged with metadata
#' using the `xlr8` system. This function reads both single-cell variables and multi-row
#' tables, guided by positional metadata produced by `summarize_metadata()`.
#'
#' @param raw_df A tibble representing the contents of an Excel workbook, typically
#'   produced by `read_excel_all()`. It must include sheet names, row numbers, and one
#'   column per Excel column (e.g., `x1`, `x2`, etc.).
#' @param all_info A tibble of metadata produced by `summarize_metadata()` (specifically
#'   the `all_info` entry). It defines the position of each variable and table/column.
#' @param fix_dates_regex Optional string specifying a regular expression used to detect
#'   column names that should be converted from Excel numeric date format. Defaults to
#'   `"date_|_date|^date$"`. Set to `NULL` to skip date conversion.
#' @param fan_info Optional tibble of fan metadata produced by `summarize_metadata()`
#'   (the `fan_info` entry). When supplied and non-empty, fan-rendered tables are
#'   reconstructed from their per-sheet instances and merged into the result with the
#'   same nested-tibble shape a flat table would produce. `NULL` (default) skips fan
#'   handling, preserving behaviour for templates without fan sheets.
#'
#' @return A tibble with a single row:
#'   - Each variable becomes a column with a single value.
#'   - Each table becomes a nested tibble-column, with its name derived from the metadata.
#'
#' @details
#' This function assumes the Excel input was generated from a locked, tagged template.
#' The metadata must indicate the location of each variable and table using `*((var((...))`,
#' `*((tbl((...))`, and `*((col((...))` tags. It reconstructs tables by matching each
#' column to its metadata-defined position and name, then combines them row-wise.
#'
#' If date columns in tables or variables are stored as numeric Excel dates,
#' they are automatically converted based on a regex match unless `fix_dates_regex` is `NULL`.
#'
#' @examples
#' \dontrun{
#' raw_df <- read_excel_all("form_filled.xlsx")
#' metadata <- summarize_metadata("template.xlsx")
#' xlr8_read(raw_df, metadata$all_info[[1]])
#' }
#'
#' @importFrom dplyr mutate select filter if_else bind_cols across everything matches where
#' @importFrom tidyr pivot_wider nest unnest
#' @importFrom purrr map map2 map_int pmap
#' @importFrom gplyr quickm filter_out_na filter_in_na parse_guess_all
#' @importFrom janitor excel_numeric_to_date
#' @importFrom magrittr extract
#'
#' @export

xlr8_read <- function(raw_df, all_info, fix_dates_regex = "date_|_date|^date$",
                      fan_info = NULL) {

  has_fans <- !is.null(fan_info) && nrow(fan_info) > 0
  all_info_full <- all_info
  if (has_fans) {
    # Fan tables are reconstructed separately from their per-sheet instances;
    # remove them from the normal (positional, single-sheet) extraction path.
    all_info <- all_info %>% filter(!(tbl %in% fan_info$tbl))
  }

  pull_table <- function(df, top, bottom, left, right){
    # Guard against an inverted range (e.g. an empty table whose end marker sits
    # above its start): base R's `top:bottom` would otherwise count downwards
    # and return reversed/garbage rows instead of an empty result.
    if (is.na(top) || is.na(bottom) || top > bottom) {
      return(df[0, left:right, drop = FALSE])
    }
    df %>%
      extract(top:bottom, left:right)
  }

  fix_dates_function <- function(df) {
    df %>%
      quickm(matches(fix_dates_regex) &
               where(is.numeric),
             excel_numeric_to_date)
  }

  collapse_into_table <- function (df) {
    df %>%
      pivot_wider(names_from = "col_name", values_from = "column") %>%
      mutate(across(everything(),
                    .fns = ~ map2(.x, cur_column(), set_names))) %>%
      unnest(everything(), keep_empty = TRUE)
  }

  # When all remaining (non-fan) metadata is empty -- e.g. a template whose only
  # content is fan sheets -- the normal positional extraction has nothing to do.
  # Skip it to avoid degenerate pivot_wider/unnest behaviour on zero rows.
  if (nrow(all_info) > 0) {
    data_by_table_and_column_name <- all_info %>%
      mutate(
        raw_df = list(raw_df),
        step_1 = map2(
          raw_df,
          sheet_name,
          ~ .x %>% filter(sheet_name == .y) %>% select(-sheet_name, -row)
        ),
        step_1_length = map_int(step_1, nrow),
        row_end = if_else(is.na(row_end), step_1_length, row_end),
        column = pmap(
          list(
            df = step_1,
            top = row_start,
            bottom = row_end,
            left = col_start,
            right = col_start
          ),
          pull_table
        )
      ) %>%
      select(tbl, col_name, column)

    tables <- data_by_table_and_column_name %>%
      filter_out_na(col_name) %>%
      nest(data = -tbl) %>%
      pivot_wider(names_from = "tbl", values_from = "data") %>%
      quickm(everything(), map, collapse_into_table) %>%
      quickm(everything(), map, parse_guess_all)

    if (!is.null(fix_dates_regex)) {
      tables <- tables %>%
        quickm(everything(), map, fix_dates_function)
    }

    vars <- data_by_table_and_column_name %>%
      filter_in_na(col_name) %>%
      select(-col_name) %>%
      quickm(column, map, unlist) %>%
      pivot_wider(names_from = "tbl", values_from = "column") %>%
      unnest(everything(), keep_empty = TRUE) %>%
      parse_guess_all()

    if (!is.null(fix_dates_regex)) {
      vars <- vars %>%
        fix_dates_function()
    }
  } else {
    vars <- tibble()
    tables <- tibble()
  }

  pieces <- list(vars, tables)

  if (has_fans) {
    fan_tables <- collapse_fan_sheets(raw_df, all_info_full, fan_info)
    if (!is.null(fix_dates_regex)) {
      fan_tables <- fan_tables %>%
        quickm(everything(), map, fix_dates_function)
    }
    pieces <- c(pieces, list(fan_tables))
  }

  # A piece with no columns (e.g. `tables` when a template has no flat-table
  # columns at all -- only vars and/or fan tables) is a literal `tibble()`
  # produced by `pivot_wider()` on zero input rows. Including it in `bind_cols()`
  # would force the whole result down to zero rows even though the other pieces
  # have one; drop such pieces instead, since they carry no information.
  pieces <- pieces[vapply(pieces, ncol, integer(1)) > 0]

  if (length(pieces) == 0) tibble() else do.call(bind_cols, pieces)
}

#' Reconstruct Fan-Rendered Tables from Their Per-Sheet Instances
#'
#' Detects the sheets in a filled workbook that are instances of a fan-rendered
#' table (by the literal `*((fan*((...` tag text that survives into each clone),
#' extracts one row of data from each, recovers the naming-column value from the
#' real sheet (tab) name, and recombines them into a nested tibble identical in
#' shape to what the same table would produce if rendered as a flat table.
#'
#' @param raw_df A tibble from `read_excel_all()` for the filled workbook.
#' @param all_info The template's full `all_info` (including fan-table rows),
#'   used for the within-sheet position of each field.
#' @param fan_info The template's `fan_info`.
#'
#' @return A one-row tibble with one list-column per fan table, each holding the
#'   reconstructed nested tibble (zero rows if no instances were found).
#'
#' @importFrom dplyr filter distinct pull select mutate
#' @importFrom tidyr pivot_longer
#' @importFrom purrr map pmap
#' @importFrom tibble as_tibble tibble
#' @importFrom gplyr filter_in_str parse_guess_all
#' @importFrom stringr str_c
#' @importFrom stats setNames
#' @keywords internal
collapse_fan_sheets <- function(raw_df, all_info, fan_info) {

  pull_cell_value <- function(step_1, top, left) {
    if (is.na(top) || is.na(left) || top < 1 || left < 1 ||
        top > nrow(step_1) || left > ncol(step_1)) {
      return(NA)
    }
    step_1[[left]][[top]]
  }

  fans <- fan_info %>% distinct(tbl, fan_tab_name_col)

  # Cells whose surviving text is a fan tag identify each fan-instance sheet.
  fan_tag_cells <- raw_df %>%
    pivot_longer(-c(row, sheet_name)) %>%
    filter_in_str(value, str_c("^", xlr8_tag), na.rm = TRUE) %>%
    mutate(fan_tbl = target_extraction(value, "fan")) %>%
    filter(!is.na(fan_tbl))

  result_cols <- list()

  for (i in seq_len(nrow(fans))) {
    this_tbl <- fans$tbl[[i]]
    naming_col <- fans$fan_tab_name_col[[i]]

    instance_sheets <- fan_tag_cells %>%
      filter(fan_tbl == this_tbl) %>%
      pull(sheet_name) %>%
      unique()

    fan_cols <- all_info %>%
      filter(tbl == this_tbl, !is.na(col_name)) %>%
      select(col_name, row_start, col_start)

    instance_rows <- map(instance_sheets, function(inst) {
      step_1 <- raw_df %>%
        filter(sheet_name == inst) %>%
        select(-sheet_name, -row)

      vals <- pmap(
        list(top = fan_cols$row_start, left = fan_cols$col_start),
        function(top, left) pull_cell_value(step_1, top, left)
      )
      names(vals) <- fan_cols$col_name

      row_tbl <- as_tibble(vals)
      # The tab name is the authoritative source for the naming column.
      row_tbl[[naming_col]] <- inst
      row_tbl
    })

    nested <- if (length(instance_rows) == 0) {
      empty_names <- union(fan_cols$col_name, naming_col)
      as_tibble(setNames(rep(list(character()), length(empty_names)), empty_names))
    } else {
      dplyr::bind_rows(instance_rows)
    }

    nested <- parse_guess_all(nested)
    result_cols[[this_tbl]] <- list(nested)
  }

  as_tibble(result_cols)
}
