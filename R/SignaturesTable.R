# Copyright (c) 2023-2024 Novartis, rdocx authors
#
# This file is part of rdocx.
#
# Licensed under the MIT license:
#
#     http://www.opensource.org/licenses/mit-license.php
#
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
#
#   The above copyright notice and this permission notice shall be included in all
# copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
# SOFTWARE.

#' @title Class that defines the Signatures Table Row in the Generic Report
#'
#' @description This R6 class controls one row of the Signatures table that documents
#' the date, version, reason and what changed in the update on the document.
#'
#' @field name Active binding for setting the name of the person who will sign
#' @field department Active binding for setting the department of the person who signs
#' @field date Active binding for setting the date of sign
#'
#' @export
#' @import R6
#'
SignaturesTableRow <- R6::R6Class(
  "SignaturesTableRow",
  private = list(
    .name = NA_character_,
    .department = NA_character_,
    .date = NA_character_
  ),
  active = list(
    name = function(value) {
      if (missing(value)) {
        private$.name
      } else {
        private$.name <- check_name(value)
        self
      }
    },
    department = function(value) {
      if (missing(value)) {
        private$.department
      } else {
        private$.department <- checkmate::assert_string(value)
        self
      }
    },
    date = function(value) {
      if (missing(value)) {
        private$.date
      } else {
        private$.date <- check_string_is_date(value)
        self
      }
    }
  ),
  
  public = list(
    
    #' @param name String. Name of the person who will sign the document
    #' @param department String. Department of the person who will sign the document
    #' @param date String. Date (DD-Mmm-YYYY) when the document will be signed
    #'
    #' @examples
    #' str <- SignaturesTableRow$new(
    #'  name = 'Harry Styles',
    #'  department = 'eCompliance',
    #'  date = '15-Feb-2024')
    #'
    initialize = function(name = NA_character_,
                          department = NA_character_,
                          date = NA_character_
    ) {
      
      private$.name <- check_name(name)
      private$.department <- checkmate::expect_string(department)
      private$.date <- check_string_is_date(date)
      
    }
  )
)

#' @title Class that defines the Signatures Table in the Generic Report
#'
#' @description This R6 class controls one row of the Signatures table that documents
#' the date, version, reason and what changed in the update on the document.
#' It uses the ChangelogTableRow class to manage the different rows in the table,
#' and control the checks
#'
#' @field rows List where new rows are added
#' @export
#' @import R6
#'

SignaturesTable <- R6::R6Class(
  "SignaturesTable",
  public = list(
    rows = list(),
    
    
    #' @description Takes an object of class SignaturesTableRow.
    #' as input and adds it to the list of rows.
    #'
    #' @export
    #' @importFrom magrittr `%>%`
    #'
    #' @param signatures_table_row_object object of class SignaturesTableRow
    #'
    #' @examples
    #' str <- SignaturesTableRow$new(
    #'  name = 'Harry Styles',
    #'  department = 'eCompliance',
    #'  date = '15-Feb-2024')
    #'
    #'
    #' signatures_table <- SignaturesTable$new()
    #' signatures_table$add_row(str)
    #'
    add_row = function(signatures_table_row_object) {
      checkmate::assert_class(signatures_table_row_object, "SignaturesTableRow")
      self$rows <- c(self$rows, list(signatures_table_row_object))
    },
    
    #' @description Create a dataframe from all the rows in the row
    #'
    #' @importFrom magrittr `%>%`
    #'
    #'
    #' @examples
    #' str <- SignaturesTableRow$new(
    #'  name = 'Harry Styles',
    #'  department = 'eCompliance',
    #'  date = '15-Feb-2024')
    #'
    #'
    #' signatures_table <- SignaturesTable$new()
    #' signatures_table$add_row(str)
    #' signatures_table$to_dataframe()
    #'
    to_dataframe = function() {
      data_frame <- do.call(rbind, lapply(self$rows, function(row) {
        data.frame(
          name = row$name,
          department = row$department,
          date = row$date
        )
      }))
      colnames(data_frame) <- c("Name",
                                "Department",
                                "Date")
      # Add the signature column
      data_frame['Signatures'] <- rep.int(" ", nrow(data_frame))
      return(data_frame)
    },
    
    #' @description Generates the signature table usin to_dataframe()
    #'
    #' @export
    #' @docType methods
    #' @importFrom magrittr `%>%`
    #'
    #'
    #' @examples
    #' str_1 <- SignaturesTableRow$new(
    #'  name = 'Harry Styles',
    #'  department = 'eCompliance',
    #'  date = '15-Feb-2024')
    #'
    #' str_2 <- SignaturesTableRow$new(
    #'  name = 'Alex Turner',
    #'  department = 'Product Development',
    #'  date = '20-Feb-2024')
    #'
    #' signatures_table <- SignaturesTable$new()
    #' signatures_table$add_row(str_1)
    #' signatures_table$add_row(str_2)
    #' signatures_table$get_table()
    #'
    
    get_table = function() {
      signature_table <- self$to_dataframe() %>%
        flextable::flextable() %>%
        table_styling()
      signature_table
    }
  )
)