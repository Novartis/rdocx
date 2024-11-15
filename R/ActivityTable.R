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

#' @title Class that defines the Activity Table in the Generic Report
#'
#' @description This R6 class relates to the activity table in the generic report 
#' that documents the person who perfomed the activity and the person who validated 
#' and the respective dates.
#' Developer note: This class is developed using private and active fields rather
#' than public fields directly which facilitates more rigid behaviour after the
#' class has been initialised. If it is deemed that edit the fields after initialisation
#' is a rare case then we might consider simplifying these classes to use public
#' fields.
#' 
#' @field activity Active binding for setting the activity
#' @field main_author_name Active binding for setting the main author name
#' @field main_activity_date Active binding for setting the main acitivty date
#' @field qc_author_name Active binding for setting the QC author name
#' @field qc_activity_date Active binding for setting the QC activity date
#'
#' @export
#' @import R6
#'
ActivityTable <- R6::R6Class(
  "ActivityTable",
  private = list(
    .activity = NA_character_,
    .main_author_name = NA_character_,
    .main_activity_date = NA_character_,
    .qc_author_name = NA_character_,
    .qc_activity_date = NA_character_
  ),
  active = list(
    activity = function(value) {
      if (missing(value)) {
        private$.activity
      } else {
        private$.activity <- checkmate::assert_string(value)
        self
      }
    },
    main_author_name = function(value) {
      if (missing(value)) {
        private$.main_author_name
      } else {
        private$.main_author_name <- check_name(value)
        self
      }
    },
    main_activity_date = function(value) {
      if (missing(value)) {
        private$.main_activity_date
      } else {
        private$.main_activity_date <- check_string_is_date(value)
        self
      }
    },
    qc_author_name = function(value) {
      if (missing(value)) {
        private$.qc_author_name
      } else {
        private$.qc_author_name <- check_name(value)
        self
      }
    },
    qc_activity_date = function(value) {
      if (missing(value)) {
        private$.qc_activity_date
      } else {
        private$.qc_activity_date <- check_string_is_date(value)
        self
      }
    }

  ),
  public = list(

    #' @param activity Name of the activity that was performed
    #' @param main_author_name Name (Firstname Lastname) of the person who performed the activity
    #' @param main_activity_date Date (DD-Mmm-YYYY) when activity was performed.
    #' @param qc_author_name Name (Firstname Lastname) of reviewer.
    #' @param qc_activity_date Date (DD-Mmm-YYYY) when QC was performed.
    #'
    #' @export
    #'
    #' @examples
    #' at <- ActivityTable$new(
    #'  activity = 'Super cool calculation',
    #'  main_author_name = 'David Bowie',
    #'  main_activity_date = '01-Oct-2024',
    #'  qc_author_name = 'Freddie Mercury',
    #'  qc_activity_date = '20-Oct-2024')
    #'
    initialize = function(activity = NA_character_,
                          main_author_name = NA_character_,
                          main_activity_date = NA_character_,
                          qc_author_name = NA_character_,
                          qc_activity_date = NA_character_
                          ) {
      private$.activity <- checkmate::assert_string(activity)
      private$.main_author_name <- check_name(main_author_name)
      private$.main_activity_date <- check_string_is_date(main_activity_date)
      private$.qc_author_name <- check_name(qc_author_name)
      private$.qc_activity_date <- check_string_is_date(qc_activity_date)
    },


    #' @description Show activity table. Takes an object of class \code{\link[rdocx]{ActivityTable}}.
    #' as input and generates Table 1 of the Generic Report Document.
    #'
    #' @export
    #' @import flextable
    #' @importFrom magrittr `%>%`
    #'
    #' @param act An object of class \code{\link[rdocx]{ActivityTable}}.
    #' @return Table 1 of the Sample Size Document.
    #'
    #' @examples
    #' at <- ActivityTable$new(
    #'  activity = 'Super cool calculation',
    #'  main_author_name = 'David Bowie',
    #'  main_activity_date = '01-Oct-2024',
    #'  qc_author_name = 'Freddie Mercury',
    #'  qc_activity_date = '20-Oct-2024')
    #'
    #' at$get_table()
    #'
    get_table = function(){
      
      activity_row_1 <- paste0(self$activity, " performed by")
      activity_row_2 <- paste0(self$activity, " reviewed by")
      
      activity_table <- data.frame(
        Activity = c(activity_row_1,
                     activity_row_2),
        Name = c(self$main_author_name, self$qc_author_name),
        Date = c(self$main_activity_date,self$qc_activity_date)
      ) %>%
        flextable::flextable() %>%
        table_styling() %>%
        flextable::set_caption(
          flextable::as_paragraph(
            flextable::as_chunk("Table 1-1. People involved in the activity performed", 
                     props = flextable::fp_text_default(font.family = "Helvetica",
                                                        italic = TRUE,
                                                        font.size = 9)
                     )
          ), 
          style = "Table Caption",
          autonum = "autonum"
        )
      activity_table
    }

  )
)
