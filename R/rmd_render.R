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

#' Rmd Sample Size renderer
#' @description Given the filename of a generic report Rmd template, it renders it.
#' It is a wrapper that hides the process of rendering, starting by getting the
#' title and body separately, getting the study code and compound name, and merging
#' both files with the correct output name.
#'
#' @param rmd_filename String. Name of the Rmd file to be rendered.
#' @param output_path String. Set the path where your generated Word file will be written to.
#' The default path will the path of the Rmd file.
#' @param version Integer. Version number of the created Word file, e.g., 0, 1, 12.
#' Will be modified to follow version numbering guidelines, e.g., _v00, _v01, v_12.
#' @param reference_docx String. Template document in .docx format providing
#' styling and structure information for rendering your docx report.
#' This is configured to be the generic report template using the function
#' \code{generic_report_template()}.
#'
#' @return Word docx file rendered from the Rmd located in `output_path`
#' @export
#' @import rmarkdown
#' @import officer
#' @import lgr
#'
#' @examples
#' \dontrun{
#' rmd_render(system.file("use_cases/01_generic_report",
#'                        "generic_report_template.Rmd",
#'                        package="rdocx"))
#' }
rmd_render <- function(rmd_filename,
                       output_path = dirname(rmd_filename),
                       version = 0,
                       reference_docx = generic_report_template()) {
  # Checks ------------------------------------------------------------------
  # Check the rmd_filename arg points to a file that exists
  checkmate::assert_file_exists(rmd_filename)
  # check that the output_path arg is a valid directory
  checkmate::assert_directory_exists(output_path)

  # Get the path of the Rmarkdown
  rmd_path <- dirname(rmd_filename)

  # Get formatted version
  doc_version <- version_number(version)

  # Render the document: produce _title.docx and _body.docx using the style in
  # reference_docx
  rmarkdown::render(rmd_filename,
                    output_file = "_body.docx",
                    output_format = word_document(toc = TRUE, reference_docx = reference_docx),
                    quiet = TRUE
  )

  # Open the title page to get the Report title, Study title and study number
  title_summary <- officer::read_docx(file.path(rmd_path, "_title.docx")) %>%
    officer::docx_summary()
  generic_report_title <- title_summary$text[which(title_summary$style_name %in% "GenericReportTitle")]
  study_title <- title_summary$text[which(title_summary$style_name %in% "StudyTitle")]

  # Note that `Dedicatednumber` corresponds to two fields: the study number and
  # study title -- the assumption here is that study number is always first
  study_number <- title_summary$text[which(title_summary$style_name %in% "StudyNumber")]

  # Create the name of the final file
  final_document_name <- paste0(
    study_number,
    "_",
    gsub(" ", "_", study_title),
    "_",
    gsub(" ", "_", generic_report_title),
    doc_version
  )

  # Merge the title and body and write the doc to `output_path`
  output <- officer::read_docx(file.path(rmd_path, "_title.docx")) %>%
    officer::body_add_docx(file.path(rmd_path, "_body.docx")) %>%
    officer::headers_replace_all_text("Generic Report Title", generic_report_title, warn = F) %>%
    officer::headers_replace_all_text("Study title", study_title, warn = F) %>%
    officer::headers_replace_all_text("Study number", study_number, warn = F) %>%
    print(target = file.path(output_path, paste0(final_document_name, ".docx")))

  # Delete the artifacts for title and body
  invisible(file.remove(file.path(rmd_path, "_title.docx")))
  invisible(file.remove(file.path(rmd_path, "_body.docx")))


  # Set up logging ----------------------------------------------------------
  log_filename <- paste0(output_path, "/", final_document_name, ".log")
  log_break <- "===================================================================================================================================="

  # Remove old logger file if necessary
  if (file.exists(log_filename)) { file.remove(log_filename) }
  # Set AppenderFile for logger
  lg_ss <- get_logger("file_ss")$set_appenders(AppenderFile$new(log_filename))

  # Set format of log file messages
  lg_ss$appenders[[1]]$layout$set_fmt("[%t]  %m")


  # Log document details and session info ---------------------------------------
  lg_ss$info("======================== Generic Report Details ============================================================================")
  lg_ss$info(paste("Document name:",  paste0(final_document_name, ".docx")))
  lg_ss$info(paste("Document location after rendering:", output_path))
  lg_ss$info(paste("Date and time:", Sys.time(), Sys.timezone()))
  lg_ss$info("")
  lg_ss$info("======================== R session info ==========================================================================================")
  for ( lib in .libPaths() ) { lg_ss$info(cat(lib)) }
  lg_ss$info("")
  lg_ss$info(capture.output(sessionInfo()))
  lg_ss$info("")
  lg_ss$info(log_break)
  lg_ss$info("Rendering completed.")

}
