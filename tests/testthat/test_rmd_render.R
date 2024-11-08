print(generic_report_template())

# Test 1: checks for the template docx for the sample size use case

test_that("The global variable defining the template exists and points
                    to docx file in the expected format", {
  expect_true(file.exists(generic_report_template()))
})


rmd_filename <- system.file("use_cases/01_generic_report/",
  "generic_report_template.Rmd",
  package = "rdocx"
)
output_path <- dirname(rmd_filename)

# Test 2: check that rendering runs successfully
test_that("`rmd_render` executes successfully using the example provided", {
  expect_no_error(rmd_render(
    rmd_filename = rmd_filename,
    output_path = output_path
  ))
  expect_true(
    file.exists(
      list.files(output_path, pattern = ".docx", full.names = TRUE)
    )
  )
  expect_true(
    file.exists(
      list.files(output_path, pattern = ".log", full.names = TRUE)
    )
  )
})

file.remove(list.files(output_path, pattern = ".docx", full.names = TRUE))
file.remove(list.files(output_path, pattern = ".log", full.names = TRUE))
