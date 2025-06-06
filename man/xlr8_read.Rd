% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/xlr8_read.R
\name{xlr8_read}
\alias{xlr8_read}
\title{Read Structured Data from Tagged Excel Forms}
\usage{
xlr8_read(raw_df, all_info, fix_dates_regex = "date_|_date|^date$")
}
\arguments{
\item{raw_df}{A tibble representing the contents of an Excel workbook, typically
produced by `read_excel_all()`. It must include sheet names, row numbers, and one
column per Excel column (e.g., `x1`, `x2`, etc.).}

\item{all_info}{A tibble of metadata produced by `summarize_metadata()` (specifically
the `all_info` entry). It defines the position of each variable and table/column.}

\item{fix_dates_regex}{Optional string specifying a regular expression used to detect
column names that should be converted from Excel numeric date format. Defaults to
`"date_|_date|^date$"`. Set to `NULL` to skip date conversion.}
}
\value{
A tibble with a single row:
  - Each variable becomes a column with a single value.
  - Each table becomes a nested tibble-column, with its name derived from the metadata.
}
\description{
Extracts structured data from Excel workbooks that have been pre-tagged with metadata
using the `xlr8` system. This function reads both single-cell variables and multi-row
tables, guided by positional metadata produced by `summarize_metadata()`.
}
\details{
This function assumes the Excel input was generated from a locked, tagged template.
The metadata must indicate the location of each variable and table using `*((var((...))`,
`*((tbl((...))`, and `*((col((...))` tags. It reconstructs tables by matching each
column to its metadata-defined position and name, then combines them row-wise.

If date columns in tables or variables are stored as numeric Excel dates,
they are automatically converted based on a regex match unless `fix_dates_regex` is `NULL`.
}
\examples{
\dontrun{
raw_df <- read_excel_all("form_filled.xlsx")
metadata <- summarize_metadata("template.xlsx")
xlr8_read(raw_df, metadata$all_info[[1]])
}

}
