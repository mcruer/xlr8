#' xlr8: Automate Excel Reports with Template-Based Data and Styling
#'
#' `xlr8` simplifies the creation of consistently styled Excel reports by:
#' - Reading metadata tags from Excel templates.
#' - Writing data into Excel according to extracted metadata.
#' - Automatically applying formatting from the template to populated data.
#'
#' @section Main functions:
#' - `xlr8()`: Main workflow orchestrator—reads metadata, populates Excel with data, applies template styles.
#' - `read_excel_all()`: Reads all cell data from Excel sheets into structured tibbles.
#' - `summarize_metadata()`: Extracts and summarizes embedded metadata from Excel templates.
#'
#' For more detailed workflows and examples, see the package vignettes and documentation for individual functions.
#'
#' @docType package
#' @name xlr8
NULL
