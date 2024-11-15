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

#' @title Class that defines the Changelog Table Row Table in the Generic Report
#'
#' @description This R6 class controls one row of the Changelog table that documents
#' the date, version, reason and what changed in the update on the document.
#'
#' @field date Active binding for setting the date of changes
#' @field version Active binding for setting the version of the document
#' @field why_update Active binding for setting the reason for the update
#' @field what_changed Active binding for setting the section and title impacted
#'
#' @export
#' @import R6
#'
ChangelogTableRow <- R6::R6Class(
  "ChangelogTableRow",
  private = list(
    .date = NA_character_,
    .version = NA_character_,
    .why_update = NA_character_,
    .what_changed = NA_character_
  ),
  active = list(
    date = function(value) {
      if (missing(value)) {
        private$.date
      } else {
        private$.date <- check_string_is_date(value)
        self
      }
    },
    version = function(value) {
      if (missing(value)) {
        private$.version
      } else {
        private$.version <- checkmate::assert_string(value)
        self
      }
    },
    why_update = function(value) {
      if (missing(value)) {
        private$.why_update
      } else {
        private$.why_update <- checkmate::expect_string(value,
                                                       pattern = ' ',
                                                       info = "Expected a sentence with spaces but
                                                         entry was a string with no spaces.")
        self
      }
    },
    what_changed = function(value) {
      if (missing(value)) {
        private$.what_changed
      } else {
        private$.what_changed <- checkmate::expect_string(value,
                                                          pattern = ' ',
                                                          info = "Expected a sentence with spaces but
                                                         entry was a string with no spaces.")
        self
      }
    }

  ),

  public = list(

    #' @param date String. Date (DD-Mmm-YYYY) when modifications were performed
    #' @param version String. Document version
    #' @param why_update String. Reason for updating the document
    #' @param what_changed String. Sections and titles that where impacted with the modifications
    #'
    #' @examples
    #' ctr <- ChangelogTableRow$new(
    #'  date = '01-Oct-2024',
    #'  version = 'Initial version',
    #'  why_update = 'Create first version of the document',
    #'  what_changed = 'Document creation')
    #'
    initialize = function(date = NA_character_,
                          version = NA_character_,
                          why_update = NA_character_,
                          what_changed = NA_character_
                          ) {
      private$.date <- check_string_is_date(date)
      private$.version <- checkmate::assert_string(version)
      private$.why_update <- checkmate::expect_string(why_update,
                                                     pattern = ' ',
                                                     info = "Expected a sentence with spaces but
                                                         entry was a string with no spaces.")
      private$.what_changed <- checkmate::expect_string(what_changed,
                                                        pattern = ' ',
                                                        info = "Expected a sentence with spaces but
                                                         entry was a string with no spaces.")

    }
  )
)

#' @title Class that defines the Changelog Table in the Generic Report
#'
#' @description This R6 class controls one row of the Changelog table that documents
#' the date, version, reason and what changed in the update on the document.
#' It uses the ChangelogTableRow class to manage the different rows in the table,
#' and control the checks
#'
#' @field rows List where new rows are added
#' @export
#' @import R6
#'

ChangelogTable <- R6::R6Class(
  "ChangelogTable",
  public = list(
    rows = list(),


    #' @description Takes an object of class ChangelogTableRow.
    #' as input and adds it to the list of rows.
    #'
    #' @export
    #' @importFrom magrittr `%>%`
    #'
    #' @param changelog_table_row_object object of class ChangelogTableRow
    #'
    #' @examples
    #' ctr <- ChangelogTableRow$new(
    #'  date = '01-Oct-2024',
    #'  version = 'Initial version',
    #'  why_update = 'Create first version of the document',
    #'  what_changed = 'Document creation')
    #'
    #' changelog_table <- ChangelogTable$new()
    #' changelog_table$add_row(ctr)
    #'
    add_row = function(changelog_table_row_object) {
      checkmate::assert_class(changelog_table_row_object, "ChangelogTableRow")
      self$rows <- c(self$rows, list(changelog_table_row_object))
    },

    #' @description Create a dataframe from all the rows in the row
    #'
    #' @importFrom magrittr `%>%`
    #'
    #'
    #' @examples
    #' ctr <- ChangelogTableRow$new(
    #'  date = '01-Oct-2024',
    #'  version = 'Initial version',
    #'  why_update = 'Create first versoin of the document',
    #'  what_changed = 'Document creation')
    #'
    #' changelog_table <- ChangelogTable$new()
    #' changelog_table$add_row(ctr)
    #' changelog_table$to_dataframe()
    #'
    to_dataframe = function() {
      data_frame <- do.call(rbind, lapply(self$rows, function(row) {
        data.frame(
          date = row$date,
          version = row$version,
          why_update = row$why_update,
          what_changed = row$what_changed
        )
      }))
      colnames(data_frame) <- c("Date",
                                "Version",
                                "Why we changed it",
                                "What changed")
      return(data_frame)
    },

    #' @description Generates the changelog table usin to_dataframe()
    #'
    #' @export
    #' @docType methods
    #' @importFrom magrittr `%>%`
    #'
    #'
    #' @examples
    #' ctr_1 <- ChangelogTableRow$new(
    #'  date = '01-Oct-2024',
    #'  version = 'Initial version',
    #'  why_update = 'Create first versoin of the document',
    #'  what_changed = 'Document creation')
    #'
    #'  ctr_2 <- ChangelogTableRow$new(
    #'  date = '15-Oct-2024',
    #'  version = 'Final version',
    #'  why_update = 'Change calculation of param 1',
    #'  what_changed = 'Section 2')
    #'
    #' changelog_table <- ChangelogTable$new()
    #' changelog_table$add_row(ctr_1)
    #' changelog_table$add_row(ctr_2)
    #' changelog_table$get_table()
    #'

    get_table = function() {
      changelog_table <- self$to_dataframe() %>%
        flextable::flextable() %>%
        table_styling()
      changelog_table
     }
  )
)


















