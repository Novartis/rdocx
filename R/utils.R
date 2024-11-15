


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

#' @title Replace text with officer
#' @description Search for a value in a document and replicate it with the new one
#'
#' @param doc Docx document to be modified
#' @param old_value Word to search in the doc
#' @param new_value Replacement for the old_value
#'
#' @import officer
#' @importFrom magrittr `%>%`
#' 
#' @keywords internal
#'
#' @examples
#' \dontrun{
#' replace_text("dummy_doc.docx", "title", "new title")
#' }
replace_text <- function(doc, old_value, new_value) {
  doc_updated <- doc %>%
    officer::cursor_reach(old_value) %>%
    officer::body_replace_all_text(
      old_value = old_value,
      new_value = new_value,
      only_at_cursor = TRUE, fixed = TRUE
    )
  return(doc_updated)
}

#' Find and delete output in Word doc
#' @description After escaping the output title or caption, it creates a regular
#' expression to find it in the text. If no occurrence is provided, it searches
#' using 'cursor_reach()'. If occurrence sis provided uses the utils function
#' 'cursor_reach_list()'
#'
#' @param template_docx Docx document to be modified
#' @param output_title Title of the output to search for
#' @param output_type Type (Table or Figure)
#' @param occurrence Int. Optional, when title is repeated is used to order them
#'
#' @return doc_updated
#'
#' @import checkmate
#' @import officer
#' @importFrom magrittr `%>%`
#' 
#' @keywords internal
#'
#' @examples
#' \dontrun{
#' find_and_delete_output(doc,
#'                         "Number of subjects",
#'                         "tbl",
#'                         1)
#' }

find_and_delete_output = function(template_docx,
                                  output_title,
                                  output_type,
                                  occurrence){

  checkmate::assert_choice(output_type,
    choices = c("TBL", "tbl", "Tbl", "FIG", "fig", "Fig")
  )

  output_title_escaped <- escape_caption(output_title)

  # Create regex depending on Table or Figure
  if (toupper(output_type) == "TBL") {
    regex_str <- paste0('^(Table).*[0-9]+.*', output_title_escaped, '$')
  } else if (toupper(output_type) == "FIG") {
    regex_str<- paste0('^(Figure).*[0-9]+.*', output_title_escaped, '$')
  }
  # If occurrence param is NA/NULL, means that caption should be unique, so we
  # look for it normally. This is because if there is more than one output with
  # the same caption, cursor_Reach() always ends up at the first one (as it always
  # starts looking from the beginning)
  if (is.null(occurrence) || is.na(occurrence)) {
    doc_updated <- template_docx %>%
      officer::cursor_reach(regex_str) %>%
      officer::cursor_forward() %>%
      officer::body_remove()
  } else {
    # If not, that means that captions are not unique
    # Using cursor_reach_list, we get a list of positions that matches the regex,
    # that is ordered.
    matches <- cursor_reach_list(template_docx, regex_str)
    # We use occurrence, to extract the 1st, 2nd, 3rd... occurrence, and move the
    # cursor to that position
    template_docx$officer_cursor$which <- matches$positions[occurrence]
    # On that position, we remove the Table/Figure
    doc_updated <- template_docx %>%
      officer::cursor_forward() %>%
      officer::body_remove()
  }

  return(doc_updated)
}


#' Add Generic Report Style to a flextable
#'
#' @param flex_table A flextable to modify style
#' @param widths If there is a specified width for the columns
#'
#' @return Flextable with Generic Report Style
#'
#' @import flextable
#' @importFrom magrittr `%>%`
#' 
#' @keywords internal
#'
#' @examples
#' \dontrun{
#' table_styling(table, widths = NA)
#' }
table_styling <- function(flex_table,
                           widths = NA) {
  widths <- eval(parse(text = widths))

  flextable::set_flextable_defaults(
    font.size = 10,
    font.family = "Helvetica",
    border.color = "black",
    text.align = "left"
  )

    
  flex_table <- flex_table %>%
    flextable::bg(i = 1, bg = "#b9dba4", part = "header") %>%  # Background color for the header row (Accent 3)
    flextable::color(i = 1, color = "black", part = "header") %>%  # White text for the header
    flextable::bold(i = 1, bold = TRUE, part = "header") %>%  # Bold text in the header
    flextable::align(i = 1, align = "left", part = "header") %>%  # Left align the header text
    flextable::border(border = fp_border(color = "#d3d3d3"), part = "all") %>%  # Black gridlines around all cells
    flextable::align(align = "left", part = "body") %>%  # Center align body text 
    flextable::fontsize(i=1, size=12, part='header') %>%
    flextable::fontsize(size=10, part='body')

  if (length(widths) == 1) { # table with equal widths
    flex_table <- flex_table %>% flextable::width(width = rep(6.3 / ncol_keys(flex_table), ncol_keys(flex_table)))
  } else { # manual widths in proportions
    if (length(widths) != ncol_keys(flex_table)) {
      cat(flex_table)
      cat(length(widths))
      cat(ncol_keys(flex_table))
      stop("The number of 'widths' does not match the number of columns in the table.")
    }
    for (i in 1:length(widths)) {
      flex_table <- flex_table %>% flextable::width(j = i, width = 6.3 * widths[i])
    }
  }
  return(flex_table)
}

#' Escape special characters in a string for regular expresion
#' @description This function escapes special characters in a given caption
#' string using the gsub function. Special characters can cause issues in certain
#' contexts, such as regular expressions or when processing text data. By escaping
#' these characters, the resulting string can be safely used in various applications.
#'
#' @param caption Caption/Title of a figure
#'
#' @details
#' The function uses the gsub function to replace special characters in the
#' input caption with their escaped versions. The regular expression pattern includes
#' a character class specifying the set of characters to be escaped.
#' The replacement string contains four backslashes followed by the matched special
#' character to ensure proper escaping.
#'
#' @return escaped caption
#' 
#' @keywords internal
#'
#' @examples
#' \dontrun{
#' caption <- "This is a figure title (an a parenthesis use!)"
#' escape_caption(caption)
#' }
escape_caption <- function(caption) {
  caption_escaped <- gsub("([-&*@{}^:=!/()%+?;'~|\\]|\\[|\\])", "\\\\\\1", caption)
  return(caption_escaped)
}


#' Change table in a docx
#'
#' @param template_doc Doc to modify
#' @param table Table to be added
#' @param widths If there are any pre-specified widths for columns in table
#' @param full_path Path tot eh folder where the csv is
#'
#' @return Updated document with new table
#' @import flextable
#' @import utils
#' @import stats
#' @importFrom magrittr `%>%`
#' 
#' @keywords internal
#'
#' @examples
#' \dontrun{
#' change_table(doc, "data.csv", widths = NA)
#' }
change_table <- function(template_doc,
                         full_path,
                         table,
                         widths = NA) {
  col_names <- utils::read.csv(file.path(full_path, table), header = FALSE, nrows = 1)

  flex_table <- utils::read.csv(file.path(full_path, table), header = FALSE, skip = 1) %>%
    stats::setNames(col_names) %>%
    flextable::flextable()

  if (is.null(widths)) {
    widths <- NA
  }

  flex_table <- table_styling(flex_table = flex_table, widths = widths)

  doc_updated <- template_doc %>%
    flextable::body_add_flextable(flex_table, pos='before')

  return(doc_updated)
}

#' Change figure in a docx
#'
#' @param template_doc Doc to modify
#' @param figure Figure to be added
#' @param full_path Path to the folder where the figrue is
#'
#' @return Updated document with new figure
#'
#' @import png
#' @import jpeg
#' @import officer
#' @import tools
#' @import magick
#' @importFrom magrittr `%>%`
#' 
#' @keywords internal
#'
#' @examples
#' \dontrun{
#' change_figure(doc, "/path/to/output", "output.png")
#' }
change_figure <- function(template_doc,
                          full_path,
                          figure) {

  image_ext <- tools::file_ext(figure)

  if (image_ext == "png") {
    fig <- png::readPNG(file.path(full_path, figure))
    dimensions <- dim(fig)
  } else if ((image_ext == "jpeg") | (image_ext == "jpg")) {
    fig <- jpeg::readJPEG(file.path(full_path, figure))
    dimensions <- dim(fig)
  } else if (image_ext == "bmp") {
    fig <- magick::image_read(file.path(full_path, figure))
    fig_info <- magick::image_info(fig)
    dimensions <- c(fig_info$width, fig_info$height)
  } else if (image_ext == "pdf") {
    fig <- magick::image_read_pdf(file.path(full_path, figure))[1]
    fig_info <- magick::image_info(fig)
    dimensions <- c(fig_info$width, fig_info$height)
  }

  doc_updated <- template_doc %>%
    officer::body_add_img(
      src = file.path(full_path, figure),
      pos = "before",
      width = 6.3,
      height = 6.3 * dimensions[1] / dimensions[2]
    )

  return(doc_updated)
}

#' Save updated document
#'
#' @param document Officer doc to be saved
#' @param doc_final_path Location to save
#'
#' @import officer
#' 
#' @keywords internal
#'
#' @examples
#' \dontrun{
#' save_updated_document(doc, "./new_doc.docx")
#' }
save_updated_document <- function(document,
                                  doc_final_path) {
  print(document, target = doc_final_path)
}


#' Version number
#'
#' @param version Number. The current version of the document
#'
#' @return formatted version "_v00", "_v02", _v13"
#'
#' @import checkmate
#' 
#' @keywords internal
#'
#' @examples
#' \dontrun{
#' version_number(5)
#' }
version_number <- function(version) {
  checkmate::assert_int(version, lower = 0, upper = 99)

  # Format the number with leading zeros and create the string
  formatted_version <- sprintf("%02d", version)
  formatted_version <- paste0("_v", formatted_version)

  return(formatted_version)
}


#' Cursor reach list
#' @description This function read a word document with xml2, looks for a pattern
#' in the text (i.e Figure/tables captions)/ Return the positions and text for this
#' search, and also possible sources path underneath the placeholder figure/tables
#'
#' @param x Document to be searched, already read by officer
#' @param keyword String. String to be found in the docx document
#'
#' @return List containing the positions where the keyword was found and the
#' corresponding text associated, as well as the positions where the possible source
#' path could be and its associated source path text
#'
#' @import xml2
#' 
#' @keywords internal
#'
#' @examples
#' \dontrun{
#' cursor_reach_list(doc, "example caption")
#' }
cursor_reach_list <- function(x, keyword) {
  # everything was inspired by officer's cursor_reach function
  # https://stackoverflow.com/questions/75743350/cursor-location-for-multiple-matches-officer-package-in-r

  # Get the nodes with text: searches like a path in document/body document/footer
  # and document/header
  nodes_with_text <- xml_find_all(x$doc_obj$get(), "/w:document/w:body/*|/w:ftr/*|/w:hdr/*")
  if (length(nodes_with_text) < 1) {
    stop("no text found in the document", call. = FALSE)
  }

  # get text and find the pattern we are looking for
  text_ <- xml2::xml_text(nodes_with_text)
  test_ <- grepl(pattern = keyword, x = text_)
  if (!any(test_)) {
    stop(keyword, " has not been found in the document",
         call. = FALSE)
  }
  # Get positions where it is a match and the associated text
  positions <- which(test_)
  text_fig_table_ <- text_[positions]

  # Look for sources under the Figure/Table. We add +2 to the positions because
  # in our use case, to reach the Source line, will have to move 2 positions: from
  # the caption to the placeholder figrue/table, and from the placeholder to the
  # source line
  positions_source <- positions + 2
  text_source_ <- text_[positions_source]

  # As not all the figure/table might have a source path, we filter for the ones
  # that start with "Source:"
  source_pattern <- "^Source:\\s*(.*)"
  test_source_ <- grepl(pattern = source_pattern, x = text_source_)
  # Keep the sources, and put a NA in the non-sources
  non_sources <- which(!test_source_)
  text_source_[non_sources] <- NA
  # Remove the "Source:" part of the string so that we keep only the path
  text_source_ <- sub(source_pattern, "\\1", text_source_)

  results <- list("positions" = positions,
                  "text" = text_fig_table_,
                  "positions_source" = positions_source,
                  "text_source"= text_source_)

  return(results)
}


#' Search for duplicates
#' @description Look for duplicates in captions titles and raise a warning if
#' the occurrence parameter is not defined.
#'
#' @param yaml_file List. Already read yaml file to be checked for duplicates
#' @import dplyr
#' 
#' @keywords internal
#'
search_for_duplicates = function(yaml_file) {

  df <- data.frame(title = character(), occurrence = numeric())

  for (entry in yaml_file) {
    # get title and occurrence
    title <- entry$title
    if ("occurrence" %in% names(entry) && !is.null(entry$occurrence)) {
      occurrence <- entry$occurrence
    } else {
      occurrence <- NA
    }
    # add to the df
    df <- rbind(df, data.frame(title = title, occurrence = occurrence))
  }
  # Check for duplications in title parameter
  duplicated_titles <- df %>%
    dplyr::count(title) %>%
    dplyr::filter(n > 1) %>%
    dplyr::pull(title)

  # if repeated check that occurrence is not NA
  for (tit in duplicated_titles) {
    dupes <- df %>%
      dplyr::filter(title == tit)
    if (any(is.na(dupes$occurrence))) {
      warning(paste("Duplicated title '", tit, "' missing occurrence parameter.\nConsider making the titles unique as unexpected behaviour can be encountered. If not, make sure to have added the param 'occurrence' in the yaml file"))
    }
  }
}


#' Extract figure and table captions
#' @description given a word document, extract all the caption for figures and
#' tables using 'cursor_reach_list()'. If the docx has a ToC, 'cursor_reach_list()'
#' will return captions from both the body of the text and the ToC. For that,
#' 'extract_figure_table_captions()' cleans the captions based on keywords
#' contained internally in the docx. ToC captions (keyword PAGEREF) but the first
#' caption one is always missed by the searching function. Body captions
#' (keyword STYLEREF or nothing) are always found correctly.
#' Therefore, if STYLEREF is present we filter for those. If there is no STYLEREF
#' but there is PAGEREF, we take the ones that are not PAGEREF. If not STYLEREF
#' or PAGEREF, no need to clean anything.
#'
#' @param doc_path String. Path to the docx document
#'
#' @import officer
#' @import stringr
#'
#' @return List figure_table_captions_sources contaning the figure_table_captions
#' and its corresponding sources paths
#' 
#' @keywords internal
#'
#' @examples
#' \dontrun{
#' extract_figure_table_captions(system.file("use_cases/02_automated_reporting",
#'                                          "Automated_Reporting_Example.docx",
#'                                          package="rdocx"))
#' }
extract_figure_table_captions = function(doc_path) {
  # Read docx
  doc <- officer::read_docx(doc_path)

  # Define regex for search
  regex_str <- "^(Figure|Table).*[0-9]+.*$"

  # Look for the regex in the document and extract positions matching
  matches <- cursor_reach_list(doc, regex_str)

  # Extract the matching expressions
  figure_table_captions <- matches$text
  sources_paths <- matches$text_source

  if (any(grepl("STYLEREF", figure_table_captions))) {
    # First, if all the formatting is still untouched, the good captions should have
    # the word STYLEREF
    pos_ <- grepl("STYLEREF", figure_table_captions)
    pos <- which(pos_)
    figure_table_captions <- figure_table_captions[pos]
    # Keep the associated sources with this captions
    sources_paths <- sources_paths[pos_]
  } else if (any(grepl("PAGEREF", figure_table_captions))){
    # It could happen then that for some reason, the format and styles have been
    # lost, so we would like to keep the captions from the Figures/Tables on
    # the text, not the ones from the ToC
    pos_ <- grepl("PAGEREF", figure_table_captions)
    pos <- which(!pos_)
    figure_table_captions <- figure_table_captions[pos]
    # Keep the associated sources with this captions
    sources_paths <- sources_paths[pos]
  }

  figure_table_captions_sources <- list("figure_table_captions" = figure_table_captions,
                                        "sources_paths" = sources_paths)


  return(figure_table_captions_sources)
}

#' Remove Figure/Table part of caption
#' @description Used to remove the first part fo the captions. It will clean
#' normal captions like Figure/Table followed by its numbering (i.e. Figure 1-1,
#' Table 2-1), and also more complex patterns created by using Novstyle such as
#'"Table  STYLEREF 1 \\s 2 SEQ Table \\* ARABIC \\s 1 1. Summary of Sample Statistics"
#'
#' @param caption String. Caption that need to be modified
#'
#' @return updated caption without Figure or Table number
#'
#' @keywords internal
#' @examples
#' \dontrun{
#' example_caption <- 
#' "Table  STYLEREF 1 \\s 2 SEQ Table \\* ARABIC \\s 1 1. Summary of Sample Statistics"
#' updated_example_caption <- remove_fig_table_part(example_caption)
#' expected_output <- "Summary of Sample Statistics"
#' }
remove_fig_table_part = function(caption) {
  # If the caption contains MERGEFORMAT, then we clean with the more complex regex
  # and if not we asume that ther eis not underlying compliacted pattern and clean
  # normally.
  if (any(grepl("STYLEREF", caption))) {
    # Regex that removes a string starting with Figure or Table and that contains
    # a MERGEFORMAT string followed by another MERGEFORMAT and number.
    pattern <- "^(Figure|Table)\\s+STYLEREF.+?\\d+\\.\\s*"
    updated_caption <- sub(pattern, "", caption)
    # In case there is an extra white space in front of the string
    updated_caption <- sub("^\\s+", "", updated_caption) 
  } else {
    # Regex that removes a string starting with Figure or Table followed by a space
    # and number hypen and number.
    pattern <- "(Figure|Table)\\s*([0-9]+[-]*[0-9]*)\\s*"
    updated_caption <- sub(pattern, "", caption)
  }
  return(updated_caption)
}


#' Create caption yaml
#' @description Given a list of captions, it creates a template yaml file with
#  a predefined structure and parameters to be set:
#'
#' - type. Indicates if it is a table (TBL) or a figure (FIG)
#'
#' - title. The title or caption of the table/figure in the template docx file
#'
#' - file. The name of the file (csv or image) to be used in the new version of the docx
#'
#' - widths. For tables only, a pre-defined column widths (useful for model parameter table, for example)
#' An example of the yaml structure:
#'
#' - occurrence. Int. To be used when captions are duplicated for more than one
#' figure/table. To set the order they appear in the word document.
#'
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
#'
#' @param figure_table_captions_sources List of the tables and figures with its captions
#' @param yaml_filename String. Output yaml filename, e.g. "/path/to/file/filename.yml"
#'
#' @import yaml
#' @import stringr
#' 
#' @keywords internal
#'
#' @examples
#' \dontrun{
#' captions <- extract_figure_table_captions(system.file("use_cases/02_automated_reporting",
#'                                          "Automated_Reporting_Example.docx",
#'                                          package="rdocx"))
#' create_yml(captions, "./example_yml.docx")
#' }
create_yml = function(figure_table_captions_sources,
                      yaml_filename) {

  yaml_output <- list()

  figure_table_captions <- figure_table_captions_sources$figure_table_captions
  sources_paths <- figure_table_captions_sources$sources_paths

  for (i in seq_along(figure_table_captions)) {
    entry <- list(
      type = ifelse(stringr::str_detect(figure_table_captions[[i]], "^Figure"), "FIG", "TBL"),
      title = toString(remove_fig_table_part(figure_table_captions[[i]])),
      file = if (!is.na(sources_paths[[i]])) sources_paths[[i]] else NULL,
      widths = NULL,
      occurrence = NULL
    )
    entry_name <- paste0("output_", i)

    yaml_output[[entry_name]] <- entry
  }
  yaml::write_yaml(yaml_output, yaml_filename)
}
