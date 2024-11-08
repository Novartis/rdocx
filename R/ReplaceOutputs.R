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

#' @title Replace Outputs in a word document
#'
#' @description This R6 class allows to, given a template word document and a list
#' of outputs, to add them at the given location in the text. It automates the
#' time-consuming task of adding and formatting tables and figures to a word document.
#' The two main inputs of the class are a template docx file, in this case a clinical
#' trial report and a yaml file that will contain the names of the table/figures to be replaced
#' and pointing to the new ones.
#' The yml file, should have predefined structure and parameters to be set:
#'
#' - type. Indicates if it is a table (TBL) or a figure (FIG)
#'
#' - title. The title or caption of the table/figure in the template docx file
#'
#' - file. The name of the file (csv or image) to be used in the new version of the docx
#'
#' - widths. For tables only, a pre-defined column widths (useful for model parameter table, for example)
#'
#' - occurrence. Int. To be used when captions are duplicated for more than one
#' figure/table. To set the order they appear in the word document.
#'
#' An example of the yaml structure:
#'
#' \preformatted{
#' output_1:
#'   type: TBL
#'   title: "Caption of the output 1"
#'   file: output_1.csv
#'   widths: ~
#'   occurrence: ~

#' output_2:
#'   type: FIG
#'   title: "Caption of the output 2"
#'   file: output_2.png
#'   }
#'
#'
#' @field template_docx_filename Active binding for setting the path to the template docx
#' @field outputs_path Active binding for setting the path to the folder with the outputs
#' @field doc_final_filename Active binding for setting the path for the updated docx file
#' @field yml_filename Active binding for setting the path to the yaml file
#'
#' @export
#' @import R6
#' @import checkmate
#'
ReplaceOutputs <- R6::R6Class(
  "ReplaceOutputs",
  private = list(
    .template_docx_filename = NA_character_,
    .outputs_path = NA_character_,
    .doc_final_filename = NA_character_,
    .yml_filename = NA_character_
  ),
  active = list(
    template_docx_filename = function(value){
      if (missing(value)){
        private$.template_docx_filename
      } else {
        checkmate::expect_file_exists(value, extension = "docx")
        private$.template_docx_filename <- value
        self
      }
    },
    outputs_path = function(value){
      if (missing(value)){
        private$.outputs_path
      } else {
        checkmate::assert_directory_exists(value)
        private$.outputs_path <- value
        self
      }
    },
    doc_final_filename = function(value){
      if (missing(value)){
        private$.doc_final_filename
      } else {
        checkmate::expect_path_for_output(value, overwrite = TRUE)
        private$.doc_final_filename <- value
        self
      }
    },
    yml_filename = function(value){
      if (missing(value)){
        private$.yml_filename
      } else {
        if (!is.na(value)){
          checkmate::expect_file_exists(value, extension = c("yml", "yaml"))
          private$.yml_filename <- value
        } else {
          private$.yml_filename <- value
        }
        self
      }
    }
  ),
  public = list(

    #' @param template_docx_filename String. File name of the template docx to be
    #' modified. For example, "/path/to/file/filename.docx"
    #' @param outputs_path String. Path to the folder with the outputs (figures
    #' and tables-csvs) that will be added to the new version of the docx
    #' @param doc_final_filename String. File name of the updated version of the
    #' docx file. For example, "/path/to/file/filename_updated.docx"
    #' @param yml_filename String. File name of the yaml file that will guide the
    #' replacement of the outputs (tables and figures) in  the updated docx. For
    #' example, "/path/to/file/filename.yml"
    #'
    #' @import checkmate
    #' @export
    #'
    #' @examples
    #' uo <- ReplaceOutputs$new(
    #'    template_docx_filename = system.file("use_cases/02_automated_reporting",
    #'                                         "Automated_Reporting_Example.docx",
    #'                                         package="rdocx"),
    #'    outputs_path = system.file("use_cases/02_automated_reporting/example_outputs",
    #'                               package="rdocx"),
    #'    doc_final_filename = "./test.docx",
    #'    yml_filename = system.file("use_cases/02_automated_reporting",
    #'                               "example_yml_1.yml",
    #'                               package="rdocx")
    #'  )
    initialize = function(template_docx_filename = NA_character_,
                          outputs_path = NA_character_,
                          doc_final_filename = NA_character_,
                          yml_filename = NA_character_) {

      # Checks
      checkmate::expect_file_exists(template_docx_filename, extension = "docx")
      private$.template_docx_filename <- template_docx_filename
      private$.outputs_path <- checkmate::assert_directory_exists(outputs_path)
      

      private$.doc_final_filename <- checkmate::assert_path_for_output(doc_final_filename, overwrite = TRUE)

      if (!is.na(yml_filename)){
        checkmate::expect_file_exists(yml_filename, extension = c("yml", "yaml"))
        private$.yml_filename <- yml_filename
      } else {
        private$.yml_filename <- yml_filename
      }
    },
    #'
    #' @description Get captions to yaml
    #' Given a docx file, it scans the document to find the Figures and Tables captions
    #' and provides a yaml file with all the figures and tables found. It also scans
    #' the document for possible "Source:" which has to be after a placeholder figure/table.
    #' If source was found, it will be added to "file" param in the yaml file.
    #'
    #' @param yml_caption_filename String. File name of the newly created yaml
    #' file that will contain the Figures and Tables of in the provided docx document.
    #' For example, "/path/to/file/filename.yml"
    #' It will follow the required format for the replacement of the outputs (tables
    #' and figures) to update docx.
    #' \preformatted{
    #' output_1:
    #'   type: TBL
    #'   title: "Caption of the output 1"
    #'   file: ~
    #'   widths: ~
    #'   occurrence: ~

    #' output_2:
    #'   type: FIG
    #'   title: "Caption of the output 2"
    #'   file: ~
    #'   widths: ~
    #'   occurrence: ~
    #'   }
    #'
    #' @export
    #'
    #' @examples
    #' uo <- ReplaceOutputs$new(
    #'    template_docx_filename = system.file("use_cases/02_automated_reporting",
    #'                                         "Automated_Reporting_Example.docx",
    #'                                         package="rdocx"),
    #'    outputs_path = system.file("use_cases/02_automated_reporting/example_outputs",
    #'                               package="rdocx"),
    #'    doc_final_filename = "./test.docx"
    #'  )
    #' uo$get_captions_to_yml(yml_caption_filename="./test_yml.yml")
    get_captions_to_yml = function(yml_caption_filename) {
      list_of_captions <- extract_figure_table_captions(private$.template_docx_filename)
      create_yml(list_of_captions, yml_caption_filename)
    },

    #'
    #' @description Update all outputs.
    #' It iterates through all the outputs in the yml file, and assigns them in the
    #' correct location in the updated docx.
    #' If no "file" param given (NA or NULL), it will skip the output.
    #'
    #' @import yaml
    #' @import officer
    #' @import tools
    #' @import lgr
    #' @export
    #'
    #' @examples

    #' uo <- ReplaceOutputs$new(
    #'    template_docx_filename = system.file("use_cases/02_automated_reporting",
    #'                                         "Automated_Reporting_Example.docx",
    #'                                         package="rdocx"),
    #'    outputs_path = system.file("use_cases/02_automated_reporting/example_outputs",
    #'                               package="rdocx"),
    #'    doc_final_filename = "./test.docx",
    #'    yml_filename = system.file("use_cases/02_automated_reporting",
    #'                               "example_yml_1.yml",
    #'                               package="rdocx")
    #'  )
    #' uo$update_all_outputs()
    update_all_outputs = function(){

      # Create path for the log file
      log_filename <- paste0(tools::file_path_sans_ext(private$.doc_final_filename), ".log")

      # Remove old logger file if necessary
      if (file.exists(log_filename)) { file.remove(log_filename) }

      # Set AppenderFile for logger
      lg_ar <- get_logger("file_ar")$set_appenders(AppenderFile$new(log_filename))

      # Set format of log file messages
      lg_ar$appenders[[1]]$layout$set_fmt("[%t]  %m")

      log_break <- "****************************************************************"

      lg_ar$info(paste("Word template filename: ", private$.template_docx_filename))
      lg_ar$info(paste("Outputs path: ", private$.outputs_path))
      lg_ar$info(paste("Yml filename: ", private$.yml_filename))
      lg_ar$info(paste("Final doc filename: ", private$.doc_final_filename))
      lg_ar$info(paste("Date and time:", Sys.time(), Sys.timezone()))
      lg_ar$info(paste("DaVinci user: ", Sys.getenv("USER")))
      lg_ar$info(log_break)

      doc <- officer::read_docx(private$.template_docx_filename)
      outputs_to_change <- yaml::read_yaml(private$.yml_filename)

      # Check if duplicates
      search_for_duplicates(outputs_to_change)

      for(output in outputs_to_change){
        if(!is.null(output$file)){
          lg_ar$info(paste("Output:", output$type, output$title))

          doc_updated <- find_and_delete_output(template_docx=doc,
                                                output_title=output$title,
                                                output_type=output$type,
                                                occurrence=output$occurrence)
          if(toupper(output$type)=="TBL"){
            doc_updated <- change_table(template_doc = doc_updated,
                                        full_path = private$.outputs_path,
                                        table=output$file,
                                        widths=output$widths)
            lg_ar$info(paste("Table updated with:",file.path(private$.outputs_path, output$file)))

          } else if (toupper(output$type)=="FIG") {
            doc_updated <- change_figure(template_doc = doc_updated,
                                         full_path = private$.outputs_path,
                                         figure = output$file)
            lg_ar$info(paste("Figure updated with:", file.path(private$.outputs_path, output$file)))
          }
        }
      }

      if (!exists('doc_updated')) {
        doc_updated <- doc
        lg_ar$info("Warning: The docx was not modified!")
      }

      save_updated_document(doc_updated, private$.doc_final_filename)
      lg_ar$info(paste0("Docx saved!! File located at: ", private$.doc_final_filename))

      lg_ar$info(log_break)

      lg_ar$info("Information about this R session")
      lg_ar$info(capture.output(.libPaths()))
      lg_ar$info(log_break)
      lg_ar$info(capture.output(sessionInfo()))

    }
  )
)
