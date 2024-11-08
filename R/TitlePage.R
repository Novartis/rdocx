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

#' @title Title page class definition
#'
#' @description This R6 class relates to the Title Page in a report that adds all
#' the required fields of the Generic Report Title Page
#' Developer note: This class is developed using private and active fields rather
#' than public fields directly which facilitates more rigid behaviour after the
#' class has been initialised. If it is deemed that edit the fields after initialisation
#' is a rare case then we might consider simplifying these classes to use public
#' fields.
#'
#' @field report_title Active binding for setting the report title
#' @field department Active binding for setting the department
#' @field study_title Active binding for setting the study title
#' @field study_number Active binding for setting the study number
#' @field status Active binding for setting the document status
#' @field version Active binding for setting the document version
#' @field date Active binding for setting the date
#' @field bus_class Active binding for setting the business classification
#' @field changelog_table Active binding for setting the changelog table
#' @field signatures_table Active binding for setting the signatures table
#'
#' @export
#' @import R6
#' @import checkmate
#'
TitlePage <- R6::R6Class(
  "TitlePage",
  private = list(
    .report_title = NA_character_,
    .department = NA_character_,
    .study_title = NA_character_,
    .study_number = NA_character_,
    .status = NA_character_,
    .version = NA_character_,
    .date = NA_character_,
    .bus_class = NA_character_,
    .changelog_table = NULL,
    .signatures_table = NULL
  ),
  active = list(
    report_title = function(value) {
      if (missing(value)) {
        private$.report_title
      } else {
        private$.report_title <- checkmate::expect_string(value,
          pattern = " ",
          info = "Expected a sentence with spaces but entry was a string with no spaces."
        )
        self
      }
    },
    department = function(value) {
      if (missing(value)) {
        private$.department
      } else {
        private$.department <- checkmate::expect_string(value)
        self
      }
    },
    study_title = function(value) {
      if (missing(value)) {
        private$.study_title
      } else {
        private$.study_title <- checkmate::expect_string(value,
          pattern = " ",
          info = "Expected a sentence with spaces but entry was a string with no spaces."
        )
        self
      }
    },
    study_number = function(value) {
      if (missing(value)) {
        private$.study_number
      } else {
        private$.study_number <- checkmate::assert_string(value)
        self
      }
    },
    status = function(value) {
      if (missing(value)) {
        private$.status
      } else {
        private$.status <- checkmate::assert_choice(value, c("Draft", "Final"))
        self
      }
    },
    version = function(value) {
      if (missing(value)) {
        private$.version
      } else {
        checkmate::assert_double(as.double(value), max.len = 1, any.missing = FALSE)
        private$.version <- checkmate::assert_string(value)
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
    },
    bus_class = function(value) {
      if (missing(value)) {
        private$.bus_class
      } else {
        private$.bus_class <- checkmate::assert_string(value)
        self
      }
    },
    changelog_table = function(value) {
      if (missing(value)) {
        private$.changelog_table
      } else {
        private$.changelog_table <- check_null_or_flextable(value)
        self
      }
    },
    signatures_table = function(value) {
      if (missing(value)) {
        private$.signatures_table
      } else {
        private$.signatures_table <- check_null_or_flextable(value)
        self
      }
    }
  ),
  public = list(

    #' @param report_title String. Report title.
    #' @param department String. Department.
    #' @param study_title  String. Study title of the document
    #' @param study_number String. Study number of the document
    #' @param status  String. Status of the document, with possible values `Draft` or `Final`
    #' @param version  String. Version of the document.
    #' @param date String. Date (DD-Mmm-YYYY) of the final document release.
    #' @param bus_class String. Business Classification.
    #' @param changelog_table Flextable. Table to be added after the title page (Optional)
    #' @param signatures_table Flextable. Table to be added after the title page (Optional)
    #'
    #' @export
    #' @import checkmate
    #' @examples
    #' \dontrun{
    #' tp = TitlePage$new(
    #'   report_title = "Super cool document",
    #'   department = "Cool department",
    #'   study_title = "Compute super cool things",
    #'   study_number = "OWNDJQW9923",
    #'   status = "Draft",
    #'   version  = "1",
    #'   date = "01-Feb-2024",
    #'   bus_class = "Very confidential")
    #'   }
    initialize = function(report_title = NA_character_,
                          department = NA_character_,
                          study_title = NA_character_,
                          study_number = NA_character_,
                          status = c("Draft", "Final"),
                          version = NA_character_,
                          date = NA_character_,
                          bus_class = NA_character_,
                          changelog_table = NULL,
                          signatures_table = NULL) {
      private$.report_title <- checkmate::expect_string(report_title,
                                                       pattern = " ",
                                                       info = "Expected a sentence with spaces but entry was a string with no spaces."
      )
      private$.department <- checkmate::expect_string(department)
      private$.study_title <- checkmate::expect_string(study_title,
                                                       pattern = " ",
                                                       info = "Expected a sentence with spaces but entry was a string with no spaces."
      )
      private$.study_number <- checkmate::assert_string(study_number)
      private$.status <- checkmate::assert_choice(status, c("Draft", "Final"))
      checkmate::assert_double(as.double(version), max.len = 1, any.missing = FALSE)
      private$.version <- checkmate::assert_string(version)
      private$.date <- check_string_is_date(date)
      private$.bus_class <- checkmate::assert_string(bus_class)
      private$.changelog_table <- check_null_or_flextable(changelog_table)
      private$.signatures_table <- check_null_or_flextable(signatures_table)
    },


    #' @description Generate title page. Takes an object of class \code{\link[rdocx]{TitlePage}}.
    #' as input and generates the title page.
    #' @aliases get_title_page
    #'
    #' @param reference_docx String. Template document in .docx format providing
    #' styling and structure information for rendering your docx report.
    #' This is configured to be Generic Report Template using the function
    #' \code{generic_report_template()}.
    #' @param output_path String. Path where the title page in docx format should
    #' be written. Default is the current working directory.
    #'
    #' @export
    #' @import officer
    #' @importFrom magrittr `%>%`
    #'
    #' @return A title page in Word format ('_generic_report_title.docx').
    #'
    #' @examples
    #' \dontrun{
    #' tp = TitlePage$new(
    #'   report_title= "Super cool document",
    #'   department = "Cool department",
    #'   study_title = "Compute super cool things",
    #'   study_number = "OWNDJQW9923",
    #'   status = "Draft",
    #'   version  = "1",
    #'   date = "01-Feb-2024",
    #'   bus_class = "Very confidential")
    #' tp$get_title_page()
    #' }
    get_title_page = function(reference_docx = generic_report_template(),
                              output_path = "."){
      #check that the `reference_docx` file exists
      checkmate::expect_file_exists(reference_docx, extension = "docx")
      #check that output path exists and is writable and allows overwriting
      checkmate::expect_path_for_output(output_path, overwrite = TRUE)

      # Read the template document
      doc_template <- officer::read_docx(reference_docx)
      doc_summary <- doc_template %>%
        officer::docx_summary()

      # We only want the title page, so we need the index for the last line in the title
      # page and the end of the doc
      index_end_page <- doc_summary[doc_summary$text=="Business Classification",]$doc_index + 1
      index_end_doc <- max(doc_summary$doc_index)

      doc_template <- doc_template %>%
        officer::cursor_begin() %>%
        officer::cursor_reach("Business Classification")

      # Remove all pages after the title page
      for (i in index_end_page:index_end_doc) {
        doc_template <- doc_template %>%
          officer::cursor_forward() %>%
          officer::body_remove()
      }

      title_page <- doc_template %>%
        replace_text("Generic Report Title", self$report_title) %>%
        replace_text("Department", self$department) %>%
        replace_text("Study title", self$study_title) %>%
        replace_text("Study number", self$study_number) %>%
        replace_text("Draft/Final", self$status) %>%
        replace_text("v0", self$version) %>%
        officer::body_replace_all_text(
          old_value = "dd-Mmm-yyyy",
          new_value = self$date,
          only_at_cursor = FALSE, fixed = TRUE
        ) %>%
        replace_text("Business Classification", self$bus_class)


      if (!is.null(private$.changelog_table)) {
        formatting <- officer::fp_text(font.size = 16,
                                       bold = TRUE,
                                       font.family = 'Helvetica')
        section_title <- officer::ftext(
          "Changelog Table - What has changed in every document version",
          formatting
        )

        title_page <- title_page %>%
          officer::body_add(officer::fpar(section_title)) %>%
          officer::body_add_par("") %>%
          flextable::body_add_flextable(private$.changelog_table) %>%
          body_add_break(pos = "after")
      }

      if (!is.null(private$.signatures_table)) {
        formatting <- officer::fp_text(font.size = 16,
                                       bold = TRUE,
                                       font.family = 'Helvetica')
        section_title <- officer::ftext(
          "Signature pages for Generic Document",
          formatting
        )

        title_page <- title_page %>%
          officer::body_add(officer::fpar(section_title)) %>%
          officer::body_add_par("") %>%
          flextable::body_add_flextable(private$.signatures_table) %>%
          body_add_break(pos = "after")
      }
      # Save title page separately
      print(title_page, target = "_title.docx")
    }
  )
)
