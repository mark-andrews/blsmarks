#' Validate Grade Data Frame
#'
#' Validates a data frame containing student grades by checking:
#' 1. Uniqueness of student IDs
#' 2. Validity of grade values (must be in UG or PG grade levels)
#' 3. Consistency of grade types (cannot mix UG-specific and PG-specific grades)
#'
#' @param grades_df A data frame containing student IDs and grades
#' @param id_col Character string specifying the column name for student IDs (default: 'id')
#' @param grade_col Character string specifying the column name for grades (default: 'grade')
#' @param return_df Logical, if TRUE returns the validated data frame (default: FALSE)
#'
#' @return If all validations pass and `return_df` is TRUE, returns a data frame with
#'         selected id and grade columns. Otherwise, returns NULL invisibly.
#' @export
#'
#' @examples
#' ## ---- Example set: Validation failures ---------------------------------
#'
#' # 1) Duplicate student IDs ----------------------------------------------
#' grades_dup <- data.frame(
#'   student_id = c("N12345", "N12345", "N67890"),
#'   grade_val  = c("1EXC", "1HIGH",  "1MID")
#' )
#' try(
#'   validate_grades_df(grades_dup,
#'                      id_col    = "student_id",
#'                      grade_col = "grade_val")
#' )
#'
#' # 2) Invalid grade values -----------------------------------------------
#' grades_bad <- data.frame(
#'   student_id = c("N12345", "N67890"),
#'   grade_val  = c("4GREAT", "1EXC")   # "4GREAT" is not a recognised grade
#' )
#' try(
#'   validate_grades_df(grades_bad,
#'                      id_col    = "student_id",
#'                      grade_col = "grade_val")
#' )
#'
#' # 3) Mixed UG-specific and PG-specific grades ---------------------------
#' grades_mixed <- data.frame(
#'   student_id = c("N12345", "N67890"),
#'   grade_val  = c("1HIGH",  "DHIGH")  # UG and PG schemes combined
#' )
#' try(
#'   validate_grades_df(grades_mixed,
#'                      id_col    = "student_id",
#'                      grade_col = "grade_val")
#' )
#'
#' # 4) Multiple simultaneous issues (handy for demos) ---------------------
#' grades_multi <- data.frame(
#'   student_id = c("N12345", "N12345", "N67890"),
#'   grade_val  = c("1HIGH",  "FAKE",   "DHIGH") # duplicate ID, invalid grade, mixed schemes
#' )
#' try(
#'   validate_grades_df(grades_multi,
#'                      id_col    = "student_id",
#'                      grade_col = "grade_val")
#' )
validate_grades_df <- function(grades_df, id_col = 'id', grade_col = 'grade', return_df = FALSE){
  # `grades_df` must be a data frame with one column indicating student IDs
  # (i.e., N numbers) and one column being student grades. It may have other
  # columns, but they are ignored.
  # The student ID column is indicated by the value, a string, of the argument `id_col`.
  # The grades column is indicated by the value, a string, of the argument `grade_col`.
  grades_df_selected <- dplyr::select(grades_df, dplyr::all_of(c(id=id_col, grade=grade_col)))

  # Checks =====================================================================
  cli::cli_h1("Running BLS Grade Entry Validation")
  # Track validation status
  all_valid <- TRUE
  validation_messages <- list()

  # Check 1: Unique student IDs ------------------------------------------------
  duplicate_ids <- grades_df_selected$id[duplicated(grades_df_selected$id)]
  if (length(duplicate_ids) == 0) {
    cli::cli_alert_success("All student IDs are unique")
  } else {
    all_valid <- FALSE
    msg <- paste("{.strong Duplicate student IDs found}:",
                 cli::cli_vec(unique(duplicate_ids), list("vec-trunc" = 5)))
    validation_messages <- c(validation_messages, msg)
    cli::cli_alert_danger("Duplicate student IDs found")
  }

  # Check 2: Valid grade values ------------------------------------------------
  # Define the grade level vectors
  UG_grade_levels <- c("ZERO", "1EXC", "1HIGH", "1MID", "1LOW", "21HIGH", "21MID", "21LOW",
                       "22HIGH", "22MID", "22LOW", "3HIGH", "3MID", "3LOW", "FMARG", "FMID", "FLOW")

  PG_grade_levels <- c("ZERO", "DEXC", "DHIGH", "DMID", "DLOW", "CHIGH", "CMID", "CLOW",
                       "PHIGH", "PMID", "PLOW", "FMARG", "FMID", "FLOW")

  # Define shared values
  shared_values <- intersect(UG_grade_levels,PG_grade_levels)

  # Check if all grades are valid (in either UG or PG)
  invalid_grades <- grades_df_selected$grade[!grades_df_selected$grade %in% c(UG_grade_levels, PG_grade_levels)]
  if (length(invalid_grades) == 0) {
    cli::cli_alert_success("All grades are valid")
  } else {
    all_valid <- FALSE
    msg <- paste("{.strong Invalid grades found}:",
                 cli::cli_vec(unique(invalid_grades), list("vec-trunc" = 5)))
    validation_messages <- c(validation_messages, msg)
    cli::cli_alert_danger(msg)
    cli::cli_ul(c(
      "Allowed UG grades: {.val {UG_grade_levels}}",
      "Allowed PG grades: {.val {PG_grade_levels}}"
    ))
  }

  # Check 3: Consistent grade types -------------------------------------------
  # Check if any UG-specific or PG-specific values exist
  ug_specific <- setdiff(UG_grade_levels, shared_values)
  pg_specific <- setdiff(PG_grade_levels, shared_values)

  has_ug_specific <- any(grades_df_selected$grade %in% ug_specific)
  has_pg_specific <- any(grades_df_selected$grade %in% pg_specific)

  # Determine if we have UG or PG grades for reporting
  grade_type <- NA
  if (has_ug_specific && !has_pg_specific) {
    grade_type <- "UG"
  } else if (!has_ug_specific && has_pg_specific) {
    grade_type <- "PG"
  } else if (!has_ug_specific && !has_pg_specific) {
    # Only shared values, can be either type
    grade_type <- "shared only"
  }

  # If vector has both UG-specific and PG-specific values, it's invalid
  if (has_ug_specific && has_pg_specific) {
    all_valid <- FALSE
    mixed_ug <- unique(grades_df_selected$grade[grades_df_selected$grade %in% ug_specific])
    mixed_pg <- unique(grades_df_selected$grade[grades_df_selected$grade %in% pg_specific])
    msg <- paste("{.strong Mixed UG and PG grades found}")
    validation_messages <- c(validation_messages, msg)
    cli::cli_alert_danger(msg)
    cli::cli_ul(c(
      "UG-specific grades present: {.val {mixed_ug}}",
      "PG-specific grades present: {.val {mixed_pg}}"
    ))
  } else {
    cli::cli_alert_success("Grade scheme is consistent ({.val {grade_type}})")
  }

  # Final summary -------------------------------------------------------------
  cli::cli_h2("Validation Summary")
  if (all_valid) {
    cli::cli_alert_success("All checks passed successfully!")
    cli::cli_alert_info("Found {nrow(grades_df_selected)} valid student records with {.val {grade_type}} grading scheme.")
    if (return_df) return(grades_df_selected)
  } else {
    cli::cli_alert_danger("Validation failed with {length(validation_messages)} issue(s):")
    cli::cli_ul(validation_messages)
    stop("Validation checks failed - please fix the issues above", call. = FALSE)
  }

  invisible(NULL)
  # ============================================================================
}


#' Enter student grades in BLS marks spreadsheet
#'
#' @param grades_df A data frame with student grades in NTU GBA abbreviated format like DHIGH, DMID, etc
#' @param bls_marks_spreadsheet_filename The filename of BLS marks spreadsheet, e.g. "PSYC40545_29120_M01_CWK_100-202425.xlsx"
#' @param id_col The column in `grades_df` that denotes the student ID (N number)
#' @param grade_col The column in `grades_df` with the students' grades
#'
#' @returns TRUE if BLS marks spreadsheet is successfully written, FALSE otherwise
#' @export
enter_bls_grades <- function(grades_df, bls_marks_spreadsheet_filename, id_col = 'id', grade_col = 'grade'){


  # Validated grades_df and return new data frame with the two main columns
  grades_df_selected <- validate_grades_df(grades_df, id_col = id_col, grade_col = grade_col, return_df = TRUE)

  marksheet_df <-
    readxl::read_excel(path = bls_marks_spreadsheet_filename, sheet = 'Grades') |>
    # Grade and Comment must be character
    # For some reason, they are read in as lgl
    dplyr::mutate(dplyr::across(c(Grade, Comment), as.character))

  # Create a temporary data frame with the grades for each student
  tmp_df <- dplyr::left_join(marksheet_df, grades_df_selected, by = c(`Student ID` = 'id'))

  # Get data frame of students in marksheet_df with existing grades
  existing_df <- dplyr::filter(tmp_df, !is.na(Grade))

  # Any conflicts? Already existing grades and new grades too?
  conflicts_df <- dplyr::filter(tmp_df, !is.na(Grade), !is.na(grade))

  # Are there students in grades_df_selected for whom no ID exists in the marksheet_df?
  new_students_df <- grades_df_selected |> dplyr::anti_join(marksheet_df, by = c('id' = 'Student ID'))

  # Loop through each row of marksheet_df
  counter <- 0
  for (i in seq(nrow(tmp_df))){

    # If we have a `Grade` value already, leave it as it is.
    # So only proceed if `Grade` is NA.
    if (is.na(tmp_df[i, 'Grade'])){

      # If we have no grade, then Grade is set to
      # ZERO and it is an NS in Comment
      if (is.na(tmp_df[i, 'grade'])) {
        tmp_df[i, 'Grade'] <- 'ZERO'
        tmp_df[i, 'Comment'] <- 'NS'
      } else {
        counter <- counter + 1
        # otherwise, `Grade` is set to `grade`
        tmp_df[i, 'Grade'] <- tmp_df[i, 'grade']
      }
    }

  }

  # Checks: -----------------------------------------------------------------
  # 1) Are all grades that were initially in marksheet_df, if any, also in tmp_df now?

  original_marksheet_grades <-
    marksheet_df |
    dplyr::select(id = `Student ID`, grade = Grade) %>%
    tidyr::drop_na()

  tmp_marksheet_grades <-
    tmp_df |
    dplyr::select(id = `Student ID`, grade = Grade) %>%
    tidyr::drop_na()

  stopifnot(
    dplyr::inner_join(original_marksheet_grades,
               tmp_marksheet_grades,
               by = 'id',
               suffix = c('_orig', '_tmp')) %>%
      dplyr::mutate(hit = grade_orig == grade_tmp) %>%
      dplyr::summarise(hit = all(hit)) %>%
      dplyr::pull(hit)
  )

  # 2) Are all grades in grades_df_selected, unless they were in marksheet already, in tmp_df?

  stopifnot(
    dplyr::inner_join(grades_df_selected,
               tmp_marksheet_grades %>%
                 dplyr::anti_join(original_marksheet_grades, by = 'id'),
               by = 'id',
               suffix = c('_grades', '_tmp')) %>%
      dplyr::mutate(hit = grade_grades == grade_tmp) %>%
      dplyr::summarise(hit = all(hit)) %>%
      dplyr::pull(hit)
  )

  # Write it ----------------------------------------------------------------

  # If checks pass, we write to file.

  # get the workbook as openxlsx object
  blsstudentr1_wb <- openxlsx::loadWorkbook(bls_marks_spreadsheet_filename)

  # write the new data frame
  openxlsx::writeData(wb = blsstudentr1_wb,
                      sheet = 'Grades',
                      # write `tmp_df` without grade
                      x = select(tmp_df, -grade),
                      startRow = 1, startCol = 'A', colNames = TRUE)

  # write it to file
  success <- openxlsx::saveWorkbook(blsstudentr1_wb,
                                    file = bls_marks_spreadsheet_filename,
                                    returnValue = TRUE,
                                    overwrite = TRUE)

  cat(
    glue::glue(
      "There are {nrow(grades_df_selected)} grades in the grades_df input.
      Of these, {counter} were entered into the BLS marks spreadsheet.
      The BLS spreadsheet contains rows for {nrow(marksheet_df)} {dplyr::if_else(nrow(marksheet_df) == 1, 'student', 'students')}.
      It already contained grades for {nrow(existing_df)} {dplyr::if_else(nrow(existing_df) == 1, 'student', 'students')}. These grades were all left unchanged.
      Of these {nrow(existing_df)} {dplyr::if_else(nrow(existing_df) == 1, 'student', 'students')}, we have new grades for {nrow(conflicts_df)} of them.

      "
    )
  )

  if (nrow(conflicts_df) > 0){
    lines <- glue_data(
      dplyr::select(conflicts_df, id = `Student ID`, existing_grade = Grade, new_grade = grade),
      "{sprintf('%-6s', id)}  | existing = {sprintf('%-6s', existing_grade)}  |  new = {sprintf('%-6s', new_grade)}"
    )

    cat(glue_collapse(lines, sep = "\n"), "\n")
  }

  cat(
    glue::glue(
      "\n      There was {nrow(new_students_df)} {dplyr::if_else(nrow(new_students_df) == 1, 'student', 'students')} for whom we have new grades but who {dplyr::if_else(nrow(new_students_df) == 1, 'is', 'are')} not listed in the BLS marksheet.

      "
    )
  )

  if (nrow(new_students_df) > 0){
    lines <- glue_data(
      new_students_df,
      "{sprintf('%-6s', id)}  |  grade = {sprintf('%-6s', grade)}"
    )

    cat(glue_collapse(lines, sep = "\n"), "\n")
  }


  success

}


#' Get path to example Excel data
#'
#' Returns the file path to example Excel data included with the package.
#'
#' @param filename Name of the Excel file in inst/extdata. Default is "PSYC20255_24424_M01_PHT_25-202425.xlsx".
#'
#' @return A character string with the file path
#' @export
#'
#' @examples
#' \dontrun{
#'   # Get the file path
#'   excel_path <- get_example_xlsx_path()
#'
#'   # Read the file with readxl
#'   if (requireNamespace("readxl", quietly = TRUE)) {
#'     data <- readxl::read_excel(excel_path)
#'     head(data)
#'   }
#' }
get_example_xlsx_path <- function(filename = "PSYC20255_24424_M01_PHT_25-202425.xlsx") {
  system.file("extdata", filename, package = "blsmarks")
}