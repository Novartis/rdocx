
# Define an example class
tp <- TitlePage$new(
  report_title= "Super cool document",
  department = "Cool department",
  study_title = "Compute super cool things",
  study_number = "OWNDJQW9923",
  status = "Draft",
  version  = "1",
  date = "01-Feb-2024",
  bus_class = "Very confidential",
  changelog_table = NULL)

# Test 0: class assigns correct values to attributes
test_that("ActivityTable returns correct values for params", {
  expect_equal(tp$report_title, "Super cool document")
  expect_equal(tp$department,"Cool department")
  expect_equal(tp$study_title, "Compute super cool things")
  expect_equal(tp$study_number,"OWNDJQW9923")
  expect_equal(tp$status, "Draft")
  expect_equal(tp$version, "1")
  expect_equal(tp$date, "01-Feb-2024")
  expect_equal(tp$bus_class, "Very confidential")
  expect_equal(tp$changelog_table, NULL)
})

# Test 1: class was initialized
test_that("`TitlePage` class is initialised correctly", {
  expect_s3_class(tp, c("TitlePage", "R6"), exact = TRUE)
})

# Test 2: report title
tp_report_title_error <- tp$clone(deep = TRUE)

test_that("`TitlePage` throws an error if `report_title` is not in the specified format (string)", {
  expect_error(tp_report_title_error$report_title <-  837432)
  expect_error(tp_report_title_error$report_title <-  "Document")
  expect_no_error(tp_report_title_error$report_title <-  "Cool Document")
  expect_equal(tp_report_title_error$report_title, "Cool Document")
})

# Test 3: department error
tp_department_error <- tp$clone(deep = TRUE)

test_that("TitlePage throws an error if the department is not a string", {
  expect_error(tp_department_error$department <-  3349855)
  expect_no_error(tp_department_error$department <-  "Department")
  expect_equal(tp_department_error$department, "Department")
  expect_no_error(tp_department_error$department <-  "Specific Department")
  expect_equal(tp_department_error$department, "Specific Department")
})

# Test 4: study title error
tp_study_title_error <- tp$clone(deep = TRUE)

test_that("`TitlePage` throws an error if `study_title` is not a string with spaces", {
  expect_error(tp_study_title_error$study_title <-  48947)
  expect_error(tp_study_title_error$study_title <- "Title_example_no_spaces")
  expect_no_error(tp_study_title_error$study_title <- "This is a title with spaces")
  expect_equal(tp_study_title_error$study_title, "This is a title with spaces")
})

# Test 5: study number error
tp_study_number_error <- tp$clone(deep = TRUE)

test_that("`TitlePage` throws an error if `study_title` is not a string with spaces", {
  expect_error(tp_study_number_error$study_number <-  48947)
  expect_no_error(tp_study_number_error$study_number <- "UQIUDEB082")
  expect_equal(tp_study_number_error$study_number, "UQIUDEB082")
})


# Test 6: status error
tp_status_error <- tp$clone(deep = TRUE)

test_that("`TitlePage` throws an error if `status` is not `Draft` or `Final`", {
  expect_error(tp_status_error$status <-  "Pending")
  expect_error(tp_status_error$status <-  37402)
  expect_error(tp_status_error$status <-  "Fnail")
  expect_no_error(tp_status_error$status <- "Draft")
  expect_equal(tp_status_error$status, "Draft")
  expect_no_error(tp_status_error$status <- "Final")
  expect_equal(tp_status_error$status, "Final")
})

# Test 5: version error
tp_version_error <- tp$clone(deep = TRUE)

test_that("`TitlePage` throws an error if `study_title` is not a string with spaces", {
  expect_error(tp_version_error$version <-  3)
  expect_no_error(tp_version_error$version <- "2")
  expect_equal(tp_version_error$version, "2")
})

# Test 8:  release_date not in format
tp_date_error <- tp$clone(deep = TRUE)

test_that("`TitlePage` throws an error if the `release_date` are not in the specified format", {
  expect_error(tp_date_error$date <-  "01-10-2023")
  expect_error(tp_date_error$date <-  "20-10-2023")
  expect_error(tp_date_error$date <-  "20-N-2023")
  expect_no_error(tp_date_error$date <-  "20-Feb-2023")
  expect_equal(tp_date_error$date, "20-Feb-2023")
})

# Test 9:  bus_class error
tp_bus_class_error <- tp$clone(deep = TRUE)

test_that("`TitlePage`` throws an error if the 'bus_class' is not a number bigger or equal to 1", {
  expect_no_error(tp_bus_class_error$bus_class <-  'Confidential')
  expect_error(tp_bus_class_error$bus_class <-  3298)
  expect_equal(tp_bus_class_error$bus_class, 'Confidential')
})

# Test 10: get title page method
test_that("`TitlePage$get_title_page`' method creates a docx correclty", {
  expect_no_error(tp$get_title_page())
  expect_true(file.exists("./_title.docx"))
  file.remove("./_title.docx")
})

test_that("`TitlePage$get_title_page` throws an error if the `reference_doxc` or `output_path` are mispecified", {
  expect_error(tp$get_title_page(reference_docx = "this_doc_does_not_exist.docx"),
                         regex = "File does not exist")
  expect_no_error(tp$get_title_page(reference_docx = system.file("use_cases/00_word_templates/generic_report/",
                                                                   "Generic_Report_Template.docx",
                                                                   package = "rdocx")))
  file.remove("./_title.docx")
  expect_error(tp$get_title_page(output_path = "path/to/nowhere/"),
                         regex = "does not exist")
})

# Test that a ChnagelogTable can be added to the TitlePage
ctr_1 <- ChangelogTableRow$new(
  date = '01-Oct-2024',
  version = 'Initial version',
  why_update = 'Create first version of the document',
  what_changed = 'NA/ Initial version')

ctr_2 <- ChangelogTableRow$new(
  date = '15-Oct-2024',
  version = 'Final version',
  why_update = 'Change calculation of param 1',
  what_changed = 'Section 2')

ct_example <- ChangelogTable$new()
ct_example$add_row(ctr_1)
ct_example$add_row(ctr_2)
ct_table <- ct_example$get_table()

test_that("`TitlePage` allows to add a ChangelogTable as a param", {
  expect_no_error(tp$changelog_table <- ct_table)
  expect_equal(tp$changelog_table, ct_table)
  expect_error(tp$changelog_table <- "table")
})

# Test that SignaturesTable can be added to the TitlePage
str_1 <- SignaturesTableRow$new(
  name = 'Harry Styles',
  department = 'eCompliance',
  date = '15-Feb-2024')

str_2 <- SignaturesTableRow$new(
  name = 'Alex Turner',
  department = 'Product Development',
  date = '20-Feb-2024')

st_example <- SignaturesTable$new()
st_example$add_row(str_1)
st_example$add_row(str_2)
st_table <- st_example$get_table()

test_that("`TitlePage` allows to add a ChangelogTable as a param", {
  expect_no_error(tp$signatures_table <- st_table)
  expect_equal(tp$signatures_table, st_table)
  expect_error(tp$signatures_table <- "table")
})


test_that("generic_report_template() loads correctly the reference doc", {
  expected_value <- system.file("use_cases/00_word_templates/generic_report",
                                "Generic_Report_Template.docx",
                                package = "rdocx")
  expect_equal(generic_report_template(), expected_value)
})

