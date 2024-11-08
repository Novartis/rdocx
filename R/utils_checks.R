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

#' @title Check date provided
#' @description Check to see if a string is a date format
#'
#' @param date_str String. Candidate string to check.
#' @param date_format String. Expected date format. Default is dd-Mmm-yyyy.
#'
#' @return the original date str if matches conditions, FALSE if not
#'
#' @examples
#' \dontrun{
#' check_string_is_date("01-Oct-2023")
#' }
check_string_is_date <- function(date_str, date_format = "%d-%b-%Y") {

  if (!is.na(as.Date(date_str, format = date_format))) {
    return(date_str)
  } else {
    stop(paste0("`Date` was not provided in the expected format (", date_format, ").",
         " For example: 01-Oct-2023"))
  }
}


#' @title Check name provided
#' @description Check to see if name has Firstname Lastname and provide a custom
#' error to guide users
#' @param name_str Name to be checked
#'
#' @return the original date str if matches conditions, FALSE if not
#' @import checkmate
#'
#' @examples
#' \dontrun{
#' check_name("Harry Styles")
#' }
check_name <- function(name_str){
  checkmate::expect_string(name_str,
                           pattern = ' ',
                           info = "`name` does not conform to the format
                               \'Firstname Lastname\'")
  return(name_str)
}

#' @title Check for NULL or flextable
#' @description Checks if the objects passed is either NULL or a flextable
#' @param variable object to check
#'
#' @return the variable if matches conditions, error if not
#'
#' @examples
#' \dontrun{
#' ft <- flextable::flextable(head(mtcars))
#' check_null_or_flextable(ft)
#' }
check_null_or_flextable <- function(variable) {

  if (is.null(variable) || inherits(variable, "flextable")) {
    return(variable)
  } else {
    stop("If you want to add a Changelog Table, it needs to be a flextable, the
         result of using ChangelofTable$get_Table()")
  }
}
