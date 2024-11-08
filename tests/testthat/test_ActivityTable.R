
# define a class
at <- ActivityTable$new(
  activity = 'Super cool calculation',
  main_author_name = 'David Bowie',
  main_activity_date = '01-Oct-2024',
  qc_author_name = 'Freddie Mercury',
  qc_activity_date = '20-Oct-2024')

# test 1: class was initialised
test_that("ActivityTable class is initialised correctly", {
  expect_s3_class(at, c("ActivityTable", "R6"), exact = TRUE)
})

# test 2a: class doesn't accept dates that are not formatted correctly
at_date_error <- at$clone(deep = TRUE)

test_that("ActivityTable throws an error if the dates are not in the specified format", {
  expect_no_error(at_date_error$main_activity_date <-  "01-Oct-2023")
  expect_equal(at_date_error$main_activity_date, "01-Oct-2023")
  expect_error(at_date_error$main_activity_date <-  "01-10-2023")
  expect_no_error(at_date_error$qc_activity_date <-  "20-Oct-2023")
  expect_equal(at_date_error$qc_activity_date, "20-Oct-2023")
  expect_error(at_date_error$qc_activity_date <-  "20-10-2023")
})

# test 3: class doesn't accept names that are not formatted correctly
at_main_author_name <- at$clone(deep = TRUE)

test_that("ActivityTable throws an error if the names are not in the format Firstname Lastname, i.e. two names with a space inbetween.", {
  expect_error(at_main_author_name$main_author_name <-  "Beyonce") #our class does should not accept single names
  expect_no_error(at_main_author_name$main_author_name <- "Sarah Michelle Geller")  #our class should accept triple names
  expect_equal(at_main_author_name$main_author_name, "Sarah Michelle Geller")
  expect_no_error(at_main_author_name$main_author_name <- "Sophie Ellis-Bexter")  #our class should accept double-barrel names
  expect_equal(at_main_author_name$main_author_name, "Sophie Ellis-Bexter")
})

# test 4: class produces a flextable object via the get_table() method
test_that("Rendered table is a flextable object", {
  expect_true(attr(at$get_table(), "class") == "flextable")
})

at_activity_name <- at$clone(deep = TRUE)
test_that("ActivityTable throws an error if activity is not a string.", {
  expect_error(at_activity_name$activity <-  39863) #our class does should not accept numbers
  expect_no_error(at_activity_name$activity <- "Calculation")  #our class should accept single words
  expect_equal(at_activity_name$activity, "Calculation")
  expect_no_error(at_activity_name$activity <- "A complex activity")  #our class should accept multiple words
  expect_equal(at_activity_name$activity, "A complex activity")
})

# Test 5: class assigns correct values to attributes
test_that("ActivityTable returns correct values for params", {
  expect_equal(at$activity, 'Super cool calculation')
  expect_equal(at$main_author_name,'David Bowie')
  expect_equal(at$main_activity_date, '01-Oct-2024')
  expect_equal(at$qc_author_name, 'Freddie Mercury')
  expect_equal(at$qc_activity_date, '20-Oct-2024')
})


