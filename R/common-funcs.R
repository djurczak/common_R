library(dplyr)
library(stringr)
library(tidyr)
library(tibble)
library(ensurer)
library(openxlsx)

#' @title normalize_column_names
#'
#' @description changes the column names from TSV/CSV into legal tidyverse
#' colnames. e.g. replaces whitespaces with "." or removes characters that are
#' not valid in tibble colnames
#' @export
#' @examples
#' tibble("Ratio M/L variability [%] mix2_shotgun" = character()") %>%
#' normalized_column_names() %>% colnames()
normalize_column_names <- function(data_) {
  ensure_that(tibble::is_tibble(data_), TRUE)

  ## convert to legit colnames
  valid_names = make.names(colnames(data_))
  colnames(data_) = gsub('([[:punct:]])\\1+', '\\1', valid_names)

  return(data_)
}

#' @title write_to_spreadsheet
#'
#' @description writes a list of named tibbles into XLSX via the openxlsx
#' package. Given names will become page labels.
#' @export
#' @examples
#' data = list(
#'    "a" = runif(10) %>% tibble::enframe(),
#'    "b" = runif(10) %>% tibble::enframe()
#' ) %>% write.to.spreadsheet("test.xlsx")

write_to_spreadsheet <- function(data, file) {
  spreadsheet <- createWorkbook()

  for(label in names(data)) {
    table = data[[label]]

    addWorksheet(spreadsheet, label)
    writeData(
      spreadsheet,
      label,
      table
    )
  }

  saveWorkbook(spreadsheet, file, overwrite = T)
}
