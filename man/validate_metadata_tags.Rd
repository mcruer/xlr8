% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/summarize_metadata.R
\name{validate_metadata_tags}
\alias{validate_metadata_tags}
\title{Validate Raw Tag Strings for Unbalanced or Unknown Keys}
\usage{
validate_metadata_tags(
  tag_values,
  allowed_keys = c("var", "tbl", "col", "formula", "table_end")
)
}
\arguments{
\item{tag_values}{A character vector of metadata tag strings (e.g., from a raw Excel tibble).}

\item{allowed_keys}{Character vector of valid tag keys (defaults to standard xlr8 tags).}
}
\value{
A list with two tibbles:
\describe{
  \item{unbalanced}{Rows with an uneven number of opening tags.}
  \item{unknown}{Rows with unknown tag keys.}
}
}
\description{
Checks if metadata tag strings have an uneven number of openers or contain unknown tag keys.
}
