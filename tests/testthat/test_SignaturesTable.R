str <- SignaturesTableRow$new(
  name = 'Harry Styles',
  department = 'eCompliance',
  date = '15-Feb-2024')

# Test 0: class assigns correct values to attributes
test_that("SignaturesTableRow returns correct values for params", {
  expect_equal(str$name, 'Harry Styles')
  expect_equal(str$department,'eCompliance')
  expect_equal(str$date, '15-Feb-2024')
})

# Test 1: class was initialized
test_that("SignaturesTableRow class is initialised correctly", {
  expect_s3_class(str, c("SignaturesTableRow", "R6"), exact = TRUE)
})

# Test 2: name error
str_name_error <- str$clone(deep = TRUE)

test_that("ChangelogTableRow throws an error if the date is not in the specified format", {
  expect_error(str_name_error$name <-  "Claudia")
  expect_no_error(str_name_error$name <-  "Sarah Michelle Geller")
  expect_equal(str_name_error$name, "Sarah Michelle Geller")
  expect_no_error(str_name_error$name <-  "Sophie Ellis-Bexter")
  expect_equal(str_name_error$name, "Sophie Ellis-Bexter")
})

# Test 3: department error
str_department_error <- str$clone(deep = TRUE)

test_that("ChangelogTableRow throws an error if the date is not in the specified format", {
  expect_error(str_department_error$department <-  387)
  expect_no_error(str_department_error$department <-  "Department")
  expect_equal(str_department_error$department, "Department")
  expect_no_error(str_department_error$department <-  "Super cool department")
  expect_equal(str_department_error$department, "Super cool department")
})


# Test 4: date error
str_date_error <- str$clone(deep = TRUE)

test_that("ChangelogTableRow throws an error if the date is not in the specified format", {
  expect_error(str_date_error$date <-  "01-10-2023")
  expect_error(str_date_error$date <-  "20-10-2023")
  expect_error(str_date_error$date <-  "20-N-2023")
  expect_no_error(str_date_error$date <-  "20-Jan-2023")
  expect_equal(str_date_error$date,  "20-Jan-2023")
})

# Define an example of ChangelogTable
st <- SignaturesTable$new()

# Test 7: class was initialized
test_that("SignaturesTable class is initialised correctly", {
  expect_s3_class(st, c("SignaturesTable", "R6"), exact = TRUE)
})

# Test 8: add_row doesn't work if the input is not from class ChangelogTableRow
st_add_row_error <- st$clone(deep = TRUE)

test_that("add_row() throws an error when the input is not from class SignaturesTableRow", {
  expect_error(st_add_row_error$add_row(7))
  expect_no_error(st_add_row_error$add_row(str))
  expect_equal(st_add_row_error$rows[[1]], str)
})

# Test 9: to_dataframe creates a dataframe and it has the same number of rows as rows added
str_2 <- SignaturesTableRow$new(
  name = 'Alex Turner',
  department = 'Product Development',
  date = '20-Feb-2024')

st_to_dataframe_error <- st$clone(deep = TRUE)
st_to_dataframe_error$add_row(str)
st_to_dataframe_error$add_row(str_2)

test_that("add_row() saves correctly second row", {
  expect_equal(st_to_dataframe_error$rows[[2]], str_2)
})

test_that("to_dataframe() return a dataframe with same number of rows as rows added", {
  expect_s3_class(st_to_dataframe_error$to_dataframe(), "data.frame" )
  expect_equal(nrow(st_to_dataframe_error$to_dataframe()),2)
})

# Test 10: get_table returns a flextable
st_get_table_error <- st$clone(deep = TRUE)
st_get_table_error$add_row(str)

test_that("get_table() return a flextable", {
  expect_s3_class(st_get_table_error$get_table(), "flextable" )
})

