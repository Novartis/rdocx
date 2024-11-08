# Define an example class
uo <- ReplaceOutputs$new(
  template_docx_filename = system.file("use_cases/02_automated_reporting",
                                       "Automated_Reporting_Example.docx",
                                       package="rdocx"),
  outputs_path = system.file("use_cases/02_automated_reporting/example_outputs",
                             package="rdocx"),
  doc_final_filename = "./test.docx",
  yml_filename = system.file("use_cases/02_automated_reporting",
                             "example_yml_1.yml",
                             package="rdocx")
)

# Test 0: class assigns correct values to attributes
test_that("ChangelogTableRow returns correct values for params", {
  expect_equal(uo$template_docx_filename, 
               system.file("use_cases/02_automated_reporting",
                           "Automated_Reporting_Example.docx",
                           package="rdocx")
               )
  expect_equal(uo$outputs_path,
               system.file("use_cases/02_automated_reporting/example_outputs",
                           package="rdocx")
               )
  expect_equal(uo$doc_final_filename, "./test.docx")
  expect_equal(uo$yml_filename, 
               system.file("use_cases/02_automated_reporting",
                           "example_yml_1.yml",
                           package="rdocx")
               )
})

# Test 1: class was initialized
test_that("`ReplaceOutputs` class is initialised correctly", {
  expect_s3_class(uo, c("ReplaceOutputs", "R6"), exact = TRUE)
})

# Test 2: template docx path error
uo_template_docx_filename_error <- uo$clone(deep = TRUE)

test_that("`ReplaceOutputs` throws an error if the `template_docx_filename` is mispecified",
          {
            expect_error(
              uo_template_docx_filename_error$template_docx_filename <-
                "example_outputs/parameter_estimates.csv"
            )
            expect_error(
              uo_template_docx_filename_error$template_docx_filename <-
                "this_doc_does_not_exist.docx"
            )
            expect_no_error(
              uo_template_docx_filename_error$template_docx_filename <-
                system.file(
                  "use_cases/02_automated_reporting",
                  "Automated_Reporting_Example_with_sources.docx",
                  package =
                    "rdocx"
                )
            )
            expect_equal(
              uo_template_docx_filename_error$template_docx_filename,
              system.file(
                "use_cases/02_automated_reporting",
                "Automated_Reporting_Example_with_sources.docx",
                package =
                  "rdocx"
              )
            )
          })

# Test 3: outputs path error
uo_outputs_path_error <- uo$clone(deep = TRUE)

test_that("`ReplaceOutputs` throws an error if the `outputs_path` is mispecified", {
  expect_error(uo_outputs_path_error$outputs_path <- "path/to/nowhere/")
  expect_no_error(uo_outputs_path_error$outputs_path <- system.file("use_cases/02_automated_reporting",
                                                                      package="rdocx"))
  expect_equal(uo_outputs_path_error$outputs_path, 
               system.file("use_cases/02_automated_reporting",
                           package="rdocx"))
  expect_error(uo_outputs_path_error$outputs_path <- NA)
})

# Test 4: doc final path error
uo_doc_final_filename_error <- uo$clone(deep = TRUE)

test_that("`ReplaceOutputs` throws an error if the `doc_final_filename` is mispecified", {
  expect_error(uo_doc_final_filename_error$doc_final_filename <- "path/to/nowhere/")
  expect_no_error(uo_doc_final_filename_error$doc_final_filename <- "./test_1.docx")
  expect_equal(uo_doc_final_filename_error$doc_final_filename, "./test_1.docx")
})

# Test 5: yml path error
uo_yml_filename_error <- uo$clone(deep = TRUE)

test_that("`ReplaceOutputs` throws an error if the `yml_filename` is mispecified", {
  expect_error(uo_yml_filename_error$yml_filename <- "example_outputs/parameter_estimates.csv")
  expect_error(uo_yml_filename_error$yml_filename <- "this_doc_does_not_exist.docx")
  expect_no_error(uo_yml_filename_error$yml_filename <- system.file("use_cases/02_automated_reporting",
                                                                     "example_yml_1.yml",
                                                                     package="rdocx"))
  expect_equal(uo_yml_filename_error$yml_filename, system.file("use_cases/02_automated_reporting",
                                                                        "example_yml_1.yml",
                                                                        package="rdocx"))
  expect_no_error(uo_yml_filename_error$yml_filename <- NA)
  expect_equal(uo_yml_filename_error$yml_filename, NA)
})


# Test 6: update_all_outputs generates correclt a docx
test_that("`ReplaceOutputs` 'update_all_outputs()' method creates a docx correclty", {
  uo$update_all_outputs()
  expect_true(file.exists("./test.docx"))
  file.remove("./test.docx")
  expect_true(file.exists("./test.log"))
  file.remove("./test.log")
})

# Test 7: get_captions_to_yml creates a yml file
uo_yml <- ReplaceOutputs$new(
  template_docx_filename = system.file("use_cases/02_automated_reporting",
                                       "Automated_Reporting_Example_with_sources.docx",
                                       package="rdocx"),
  outputs_path = system.file("use_cases/02_automated_reporting/example_outputs",
                             package="rdocx"),
  doc_final_filename = "./test.docx"
)

test_that("`ReplaceOutputs` 'get_captions_to_yml()' method creates a yml correclty", {
  uo_yml$get_captions_to_yml("./test.yml")
  expect_true(file.exists("./test.yml"))

})
# Test 8: get_captions_to_yml produces expected yml file
test_that("`ReplaceOutputs` 'get_captions_to_yml()' method creates a yml correclty", {
  uo_yml$get_captions_to_yml("./test.yml")
  rdocx_yml <- yaml::read_yaml("./test.yml")
  comparison_yml <- yaml::read_yaml("test_get_captions_to_yml.yml")
  expect_equal(rdocx_yml, comparison_yml)
  file.remove("./test.yml")
})


uo_2 <- ReplaceOutputs$new(
   template_docx_filename = system.file("use_cases/02_automated_reporting",
                                        "Automated_Reporting_Example.docx",
                                        package="rdocx"),
   outputs_path = system.file("use_cases/02_automated_reporting/example_outputs",
                              package="rdocx"),
   doc_final_filename = "./test.docx",
   yml_filename = "test_unmodified_doc.yml"
 )

test_that("`ReplaceOutputs` produces warning if the document is not updated", {
   expect_output(
     uo_2$update_all_outputs(),
     regexp = "Warning: The docx was not modified!"
   )
 })

 uo_3 <- ReplaceOutputs$new(
   template_docx_filename = system.file("use_cases/02_automated_reporting",
                                        "Automated_Reporting_Example.docx",
                                        package="rdocx"),
   outputs_path = system.file("use_cases/02_automated_reporting/example_outputs",
                              package="rdocx"),
   doc_final_filename = "./test.docx",
   yml_filename = "test_duplicated_fig_names.yml"
 )

test_that("`ReplaceOutputs` produces warning if duplicated fig names in yml", {
   expect_warning(
     uo_3$update_all_outputs(),
     regexp = "Duplicated title"
   )
 })

file.remove(list.files("./", pattern = "test.log", full.names = TRUE))
file.remove(list.files("./", pattern = "test.docx", full.names = TRUE))
