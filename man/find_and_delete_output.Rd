% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/utils.R
\name{find_and_delete_output}
\alias{find_and_delete_output}
\title{Find and delete output in Word doc}
\usage{
find_and_delete_output(template_docx, output_title, output_type, occurrence)
}
\arguments{
\item{template_docx}{Docx document to be modified}

\item{output_title}{Title of the output to search for}

\item{output_type}{Type (Table or Figure)}

\item{occurrence}{Int. Optional, when title is repeated is used to order them}
}
\value{
doc_updated
}
\description{
After escaping the output title or caption, it creates a regular
expression to find it in the text. If no occurrence is provided, it searches
using 'cursor_reach()'. If occurrence sis provided uses the utils function
'cursor_reach_list()'
}
\examples{
\dontrun{
find_and_delete_output(doc,
                        "Number of subjects",
                        "tbl",
                        1)
}
}
\keyword{internal}
