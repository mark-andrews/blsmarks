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

  # `grades_df` must be a data frame with one column indicating student IDs
  # (i.e., N numbers) and one column being student grades. It may have other
  # columns, but they are ignored.
  # The student ID column is indicated by the value, a string, of the argument `id_col`.
  # The grades column is indicated by the value, a string, of the argument `grade_col`.
  grades_df_selected <- dplyr::select(grades_df, dplyr::all_of(c(id=id_col, grade=grade_col)))

  # Checks =====================================================================
  # 1) Each student ID should occur exactly once
  stopifnot(length(grades_df_selected$id) == length(unique(grades_df_selected$id)))

  # 2) Each grade should be from one or other of these lists
  UG_grade_levels <- c("1EXC", "1HIGH", "1MID", "1LOW", "21HIGH", "21MID", "21LOW",
                       "22HIGH","22MID","22LOW","3HIGH", "3MID","3LOW", "FMARG", "FMID", "FLOW"
  )

  PG_grade_levels <- c('DEXC','DHIGH','DMID','DLOW','CHIGH','CMID','CLOW',
                       'PHIGH','PMID','PLOW','FMARG','FMID', 'FLOW')

  stopifnot(
    dplyr::pull(
      dplyr::summarise(grades_df_selected, result = all(grade %in% c(UG_grade_levels, PG_grade_levels))),
    result)
  )

  # 3) And if any grade is a UG grade, they all must be, and if any grade is PG, they all must be
  all_grades_are_ug <-  dplyr::summarise(grades_df_selected, result = all(grade %in% UG_grade_levels))
  all_grades_are_pg <-  dplyr::summarise(grades_df_selected, result = all(grade %in% PG_grade_levels))

  stopifnot(xor(all_grades_are_ug, all_grades_are_pg))

  # ============================================================================

  marksheet_df <-
    readxl::read_excel(path = bls_marks_spreadsheet_filename, sheet = 'Grades') |>
    # Grade and Comment must be character
    # For some reason, they are read in as lgl
    dplyr::mutate(dplyr::across(c(Grade, Comment), as.character))

  # Create a temporary data frame with the grades for each student
  tmp_df <- dplyr::left_join(marksheet_df, grades_df_selected, by = c(`Student ID` = 'id'))

  # Loop through each row of marksheet_df
  for (i in seq(nrow(tmp_df))){

    # If we have a `Grade` value already, leave it as it is.
    # So only proceed if `Grade` is NA.
    if (is.na(tmp_df[i, 'Grade'])){

      # If we have no new_grade value, then Grade is set to
      # ZERO and it is an NS in Comment
      if (is.na(tmp_df[i, 'grade'])) {
        tmp_df[i, 'Grade'] <- 'ZERO'
        tmp_df[i, 'Comment'] <- 'NS'
      } else {
        # otherwise, `Grade` is set to `grade`
        tmp_df[i, 'Grade'] <- tmp_df[i, 'grade']
      }
    }

  }

  # Checks: -----------------------------------------------------------------
  # 1) Are all grades that were initially in marksheet_df, if any, also in tmp_df now?

  original_marksheet_grades <-
    marksheet_df |>
    dplyr::select(id = `Student ID`, grade = Grade) %>%
    tidyr::drop_na()

  tmp_marksheet_grades <-
    tmp_df |>
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
  openxlsx::saveWorkbook(blsstudentr1_wb,
                         file = bls_marks_spreadsheet_filename,
                         returnValue = TRUE,
                         overwrite = TRUE)

}

