# Define an example of ChangelogTableRow
ctr <- ChangelogTableRow$new(
  date = '01-Oct-2024',
  version = 'Initial version',
  why_update = 'Create first version of the document',
  what_changed = 'NA/ Initial version')

# Test 0: class assigns correct values to attributes
test_that("ChangelogTableRow returns correct values for params", {
  expect_equal(ctr$date, '01-Oct-2024')
  expect_equal(ctr$version,'Initial version')
  expect_equal(ctr$why_update, 'Create first version of the document')
  expect_equal(ctr$what_changed, 'NA/ Initial version')
})

# Test 1: class was initialized
test_that("ChangelogTableRow class is initialised correctly", {
  expect_s3_class(ctr, c("ChangelogTableRow", "R6"), exact = TRUE)
})

# Test 2: date error
ctr_date_error <- ctr$clone(deep = TRUE)

test_that("ChangelogTableRow throws an error if the date is not in the specified format", {
  expect_error(ctr_date_error$date <-  "01-10-2023")
  expect_error(ctr_date_error$date <-  "20-10-2023")
  expect_error(ctr_date_error$date <-  "20-N-2023")
  expect_no_error(ctr_date_error$date <-  "20-Jan-2023")
  expect_equal(ctr_date_error$date, "20-Jan-2023")
})

# Test 3: time point is a string
ctr_versiont_error <- ctr$clone(deep = TRUE)

test_that("ChangelogTableRow throws an error if the version is not a string", {
  expect_error(ctr_versiont_error$version <-  3349855)
  expect_no_error(ctr_versiont_error$version <-  "Start protocol")
  expect_equal(ctr_versiont_error$version, "Start protocol")
})

# Test 4 : why_update has more than one word
ctr_why_update_error <- ctr$clone(deep = TRUE)

test_that("ChangelogTableRow throws an error if the why_update is not a string with more than 1 word", {
  expect_error(ctr_why_update_error$why_update <-  "TBD")
  expect_no_error(ctr_why_update_error$why_update <-  "Modify table")
  expect_equal(ctr_why_update_error$why_update, "Modify table")
})

# Test 5 : outcome update has more than one word
ctr_what_changed_error <- ctr$clone(deep = TRUE)

test_that("ChangelogTableRow throws an error if the what_changed is not a string with more than 1 word", {
  expect_error(ctr_what_changed_error$what_changed <-  "TBD")
  expect_no_error(ctr_what_changed_error$what_changed <-  "Updated table")
  expect_equal(ctr_what_changed_error$what_changed, "Updated table")
})


# Define an example of ChangelogTable
ct <- ChangelogTable$new()

# Test 7: class was initialized
test_that("ChangelogTable class is initialised correctly", {
  expect_s3_class(ct, c("ChangelogTable", "R6"), exact = TRUE)
})

# Test 8: add_row doesn't work if the input is not from class ChangelogTableRow
ct_add_row_error <- ct$clone(deep = TRUE)

test_that("add_row() throws an error when the input is not from class ChangelogTableRow", {
  expect_error(ct_add_row_error$add_row(7))
  expect_no_error(ct_add_row_error$add_row(ctr))
  expect_equal(ct_add_row_error$rows[[1]], ctr)
})

# Test 9: to_dataframe creates a dataframe and it has the same number of rows as rows added
ctr_2 <- ChangelogTableRow$new(
  date = '15-Oct-2024',
  version = 'Final version',
  why_update = 'Change calculation of param 1',
  what_changed = 'Section 2')

ct_to_dataframe_error <- ct$clone(deep = TRUE)
ct_to_dataframe_error$add_row(ctr)
ct_to_dataframe_error$add_row(ctr_2)

test_that("add_row() saves correctly second row", {
  expect_equal(ct_to_dataframe_error$rows[[2]], ctr_2)
})

test_that("to_dataframe() return a dataframe with same number of rows as rows added", {
  expect_s3_class(ct_to_dataframe_error$to_dataframe(), "data.frame" )
  expect_equal(nrow(ct_to_dataframe_error$to_dataframe()), 2)
})

# Test 10: get_table returns a flextable
ct_get_table_error <- ct$clone(deep = TRUE)
ct_get_table_error$add_row(ctr)

test_that("get_table() return a flextable", {
  expect_s3_class(ct_to_dataframe_error$get_table(), "flextable" )
})

