#' Capture a Worksheet's Data Validation and Conditional Formatting
#'
#' Best-effort capture of a template worksheet's data-validation (dropdown) and
#' conditional-formatting rules, so they can be re-applied to clones produced by
#' [openxlsx2::wb_clone_worksheet()]. Some `openxlsx2` versions do not carry these
#' rules across a clone reliably; capturing them here lets `expand_fan_sheets()`
#' restore them on each clone.
#'
#' @section Verification note:
#' This reaches into `openxlsx2`'s internal worksheet structure, which is not a
#' stable public API and can change between versions. It is deliberately wrapped
#' so that a structural mismatch warns rather than aborts a write. Whether
#' cloning preserves these rules on its own (making re-application a no-op) or
#' drops them (making it load-bearing) must be confirmed empirically against the
#' installed `openxlsx2` version. Data validation (dropdowns) is the higher
#' priority of the two to preserve.
#'
#' @param wb A `wbWorkbook` object.
#' @param sheet The template worksheet name to capture from.
#'
#' @return A list with elements `data_validation` and `conditional_formatting`,
#'   each either the captured internal structure or `NULL` if absent/unavailable.
#'
#' @keywords internal
capture_template_validations <- function(wb, sheet) {
  out <- list(data_validation = NULL, conditional_formatting = NULL)
  tryCatch({
    idx <- fan_sheet_index(wb, sheet)
    if (!is.na(idx)) {
      ws <- wb$worksheets[[idx]]
      if (!is.null(ws$dataValidations) &&
          length(ws$dataValidations) > 0) {
        out$data_validation <- ws$dataValidations
      }
      if (!is.null(ws$conditionalFormatting) &&
          length(ws$conditionalFormatting) > 0) {
        out$conditional_formatting <- ws$conditionalFormatting
      }
    }
  }, error = function(e) {
    warning(
      "Could not capture data validation / conditional formatting from sheet '",
      sheet, "': ", conditionalMessage(e),
      ". Cloned fan sheets may be missing dropdowns or conditional formatting.",
      call. = FALSE
    )
  })
  out
}

#' Re-Apply Captured Validation/Conditional Formatting to a Cloned Sheet
#'
#' Restores rules captured by [capture_template_validations()] onto a clone, but
#' only for rule types the clone does not already carry (so that a clone which
#' preserved them is left untouched and rules are never duplicated).
#'
#' @inheritSection capture_template_validations Verification note
#'
#' @param wb A `wbWorkbook` object.
#' @param sheet The cloned worksheet name to restore rules onto.
#' @param captured The list returned by [capture_template_validations()].
#'
#' @return The (possibly modified) `wbWorkbook` object.
#'
#' @keywords internal
reapply_validations <- function(wb, sheet, captured) {
  if (is.null(captured) ||
      (is.null(captured$data_validation) &&
       is.null(captured$conditional_formatting))) {
    return(wb)
  }
  tryCatch({
    idx <- fan_sheet_index(wb, sheet)
    if (!is.na(idx)) {
      ws <- wb$worksheets[[idx]]
      if (!is.null(captured$data_validation) &&
          (is.null(ws$dataValidations) || length(ws$dataValidations) == 0)) {
        wb$worksheets[[idx]]$dataValidations <- captured$data_validation
      }
      if (!is.null(captured$conditional_formatting) &&
          (is.null(ws$conditionalFormatting) ||
           length(ws$conditionalFormatting) == 0)) {
        wb$worksheets[[idx]]$conditionalFormatting <- captured$conditional_formatting
      }
    }
  }, error = function(e) {
    warning(
      "Could not re-apply data validation / conditional formatting to sheet '",
      sheet, "': ", conditionalMessage(e),
      ". This clone may be missing dropdowns or conditional formatting.",
      call. = FALSE
    )
  })
  wb
}

#' Index of a Sheet Within a Workbook's Worksheet List
#'
#' @param wb A `wbWorkbook` object.
#' @param sheet A sheet name.
#' @return Integer position in `wb$worksheets`, or `NA_integer_` if not found.
#' @keywords internal
fan_sheet_index <- function(wb, sheet) {
  nms <- openxlsx2::wb_get_sheet_names(wb)
  idx <- which(unname(nms) == sheet)
  if (length(idx) == 0) {
    idx <- which(names(nms) == sheet)
  }
  if (length(idx) == 0) NA_integer_ else idx[[1]]
}

#' Validate Proposed Excel Sheet (Tab) Names
#'
#' Checks a vector of proposed sheet names against Excel's naming rules without
#' silently altering them. Returns a tibble of problems (empty if all are valid).
#'
#' @param tab_names Character vector of proposed sheet names.
#'
#' @return A tibble with columns `tab_name`, `position`, and `issue`; zero rows
#'   if every name is valid.
#'
#' @importFrom tibble tibble
#' @importFrom dplyr mutate case_when
#' @importFrom stringr str_detect
#' @importFrom gplyr filter_out_na
#' @keywords internal
validate_excel_sheet_names <- function(tab_names) {
  tibble(
    tab_name = as.character(tab_names),
    position = seq_along(tab_names)
  ) %>%
    mutate(
      issue = case_when(
        is.na(tab_name) | tab_name == "" ~ "empty or NA",
        nchar(tab_name) > 31 ~ "longer than 31 characters",
        str_detect(tab_name, "[:\\\\/?*\\[\\]]") ~
          "contains an illegal character (one of : \\ / ? * [ ])",
        TRUE ~ NA_character_
      )
    ) %>%
    filter_out_na(issue)
}

#' Expand Fan-Tagged Template Sheets into One Sheet Per Data Row
#'
#' For each table declared as a "fan" in the template metadata, clones the fan's
#' template sheet once per row of that table's data, populates each clone with a
#' single row (reusing the ordinary [write_data()] / [apply_styles()] /
#' [add_formulas()] pipeline), names each clone from the row's naming-column
#' value, and removes the original template sheet from the workbook.
#'
#' Fan rows are stripped from the returned `all_info` so the caller's normal
#' write pipeline runs only against non-fan metadata.
#'
#' @param wb A `wbWorkbook` (typically from [openxlsx2::wb_load()]).
#' @param df The data tibble. Each fan table must be present as a nested
#'   list-column named after the table, exactly as for a flat table.
#' @param all_info Metadata `all_info` from [summarize_metadata()].
#' @param fan_info Metadata `fan_info` from [summarize_metadata()] (one row per
#'   fan-tagged cell). `NULL` or zero rows means there are no fans (no-op).
#'
#' @return A list with `wb` (the expanded workbook) and `all_info` (with fan-table
#'   rows removed).
#'
#' @importFrom dplyr filter distinct mutate pull
#' @importFrom gplyr pull_cell
#' @importFrom tibble tibble
#' @importFrom rlang :=
#' @export
expand_fan_sheets <- function(wb, df, all_info, fan_info) {

  if (is.null(fan_info) || nrow(fan_info) == 0) {
    return(list(wb = wb, all_info = all_info))
  }

  # One logical fan per (sheet, table, naming column). A validated template has
  # exactly one fan tag per sheet; distinct() guards against accidental repeats.
  fans <- fan_info %>%
    distinct(sheet_name, tbl, fan_tab_name_col)

  # ---- Planning phase: compute and globally validate every tab name first, so
  # the whole operation fails before the workbook is mutated. ----
  existing_sheets <- as.character(openxlsx2::wb_get_sheet_names(wb))
  fan_template_sheets <- unique(fans$sheet_name)
  non_fan_existing <- setdiff(existing_sheets, fan_template_sheets)

  plans <- list()
  for (i in seq_len(nrow(fans))) {
    template_sheet <- fans$sheet_name[[i]]
    this_tbl <- fans$tbl[[i]]
    naming_col <- fans$fan_tab_name_col[[i]]

    if (!(this_tbl %in% names(df))) {
      stop("Fan table '", this_tbl, "' (sheet '", template_sheet,
           "') is not present as a column in `df`.", call. = FALSE)
    }

    table_df <- df %>% pull_cell(this_tbl)

    if (!is.data.frame(table_df)) {
      stop("Fan table '", this_tbl, "' must be a nested tibble in `df`.",
           call. = FALSE)
    }

    if (nrow(table_df) > 0 && !(naming_col %in% names(table_df))) {
      stop("Fan naming column '", naming_col, "' for table '", this_tbl,
           "' is not present in the data for that table.", call. = FALSE)
    }

    tab_names <- if (nrow(table_df) == 0) {
      character()
    } else {
      as.character(table_df[[naming_col]])
    }

    plans[[i]] <- list(
      template_sheet = template_sheet,
      tbl = this_tbl,
      naming_col = naming_col,
      table_df = table_df,
      tab_names = tab_names
    )
  }

  all_generated <- unlist(lapply(plans, function(p) p$tab_names), use.names = FALSE)

  # Per-name Excel legality.
  bad_names <- validate_excel_sheet_names(all_generated)
  if (nrow(bad_names) > 0) {
    print(bad_names)
    stop("One or more fan tab names are not valid Excel sheet names ",
         "(see tibble above). No silent truncation is performed in this version.",
         call. = FALSE)
  }

  # Duplicates among generated names.
  dup_generated <- unique(all_generated[duplicated(all_generated)])
  if (length(dup_generated) > 0) {
    stop("Fan expansion would create duplicate sheet names: ",
         paste(dup_generated, collapse = ", "),
         ". Each fanned row must produce a unique tab name.", call. = FALSE)
  }

  # Collisions with sheets that will remain in the final workbook.
  clash_keep <- intersect(all_generated, non_fan_existing)
  if (length(clash_keep) > 0) {
    stop("Fan tab name(s) collide with existing non-fan sheet(s): ",
         paste(clash_keep, collapse = ", "), ".", call. = FALSE)
  }

  # Collisions with a fan template sheet that still physically exists in the
  # workbook at clone time (it is removed only after its clones are created).
  clash_template <- intersect(all_generated, fan_template_sheets)
  if (length(clash_template) > 0) {
    stop("Fan tab name(s) collide with a template sheet name still present in ",
         "the workbook during expansion: ",
         paste(clash_template, collapse = ", "),
         ". Rename the template sheet (its name is not used for data) or the ",
         "colliding naming-column value(s).", call. = FALSE)
  }

  # ---- Execution phase. ----
  for (p in plans) {
    template_sheet <- p$template_sheet
    this_tbl <- p$tbl
    table_df <- p$table_df
    tab_names <- p$tab_names

    if (nrow(table_df) == 0) {
      warning("Fan table '", this_tbl, "' (sheet '", template_sheet,
              "') has zero rows; no sheets were produced and the template ",
              "sheet was removed.", call. = FALSE)
      wb <- openxlsx2::wb_remove_worksheet(wb, sheet = template_sheet)
      next
    }

    fan_all_info_template <- all_info %>% filter(tbl == this_tbl)
    captured <- capture_template_validations(wb, template_sheet)

    for (r in seq_len(nrow(table_df))) {
      new_name <- tab_names[[r]]

      wb <- openxlsx2::wb_clone_worksheet(wb, old = template_sheet, new = new_name)
      wb <- reapply_validations(wb, new_name, captured)

      fan_all_info <- fan_all_info_template %>%
        mutate(sheet_name = new_name)

      fan_df <- tibble(!!this_tbl := list(table_df[r, , drop = FALSE]))

      wb <- wb %>%
        write_data(fan_df, fan_all_info) %>%
        apply_styles(fan_df, fan_all_info) %>%
        add_formulas(fan_df, fan_all_info)
    }

    wb <- openxlsx2::wb_remove_worksheet(wb, sheet = template_sheet)
  }

  list(
    wb = wb,
    all_info = all_info %>% filter(!(tbl %in% fans$tbl))
  )
}
