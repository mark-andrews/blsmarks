## Code to prepare test data frames with student IDs and grades
#### Code to prepare test data frames with student IDs and grades
##

set.seed(42)  # For reproducibility

# Define parameters
n_students <- 50
n_duplicates <- 3
id_prefix <- "S"
id_digits <- 7

UG_grade_levels <- c("ZERO", "1EXC", "1HIGH", "1MID", "1LOW", "21HIGH", "21MID", "21LOW",
                     "22HIGH","22MID","22LOW","3HIGH", "3MID","3LOW", "FMARG", "FMID", "FLOW")

PG_grade_levels <- c("ZERO", 'DEXC','DHIGH','DMID','DLOW','CHIGH','CMID','CLOW',
                     'PHIGH','PMID','PLOW','FMARG','FMID', 'FLOW')

# Generate student IDs (S followed by id_digits digits)
create_ids <- function(n_students){
  paste0(id_prefix,
         sprintf(paste0("%0", id_digits, "d"), # pad with zero if necessary
                 sample(seq(10^id_digits - 1), n_students)))
}

# Create data frames

# valid UG data set
student_grades_df1 <- tibble::tibble(
  id = create_ids(n_students),
  grade = sample(UG_grade_levels, n_students, replace = TRUE)
)

# valid PG data set
student_grades_df2 <- tibble::tibble(
  id = create_ids(n_students),
  grade = sample(PG_grade_levels, n_students, replace = TRUE)
)

# UG data set with duplicates
student_grades_df3 <-
  rbind(
    dplyr::mutate(head(student_grades_df1, n_duplicates), grade = sample(UG_grade_levels, n_duplicates, replace = TRUE)),
    head(student_grades_df1, nrow(student_grades_df1) - n_duplicates)
  ) |> dplyr::sample_frac(1)

# Write data --------------------------------------------------------------

usethis::use_data(student_grades_df1, overwrite = TRUE)
usethis::use_data(student_grades_df2, overwrite = TRUE)
usethis::use_data(student_grades_df3, overwrite = TRUE)


