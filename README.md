
<!-- README.md is generated from README.Rmd. Please edit that file -->

# blsmarks

<!-- badges: start -->
<!-- badges: end -->

The goal of blsmarks is to automate the process of entering grades into
NTU BLS marks spreadsheets.

## Installation

This package is probably never going to be on CRAN because it is just
something for people in the Psychology department in NTU. As such, it
should be installed from GitHub. There are multiple ways of doing this,
but one common way is as follows:

``` r
remotes::install_github("mark-andrews/blsmarks")
```

## How to use

Here is an example of how it can be used. Let’s assume you have a data
frame in R, let’s call it `grades_df`, that has one column named `nid`
that indicates the student IDs (N numbers) and another column named
`grade` with student grades (in their GBA abbreviated forms like
“21HIGH”, “DMID” etc). And let’s assume we have a BLS marks Excel
spreadsheet like “PSYC40545_29120_M01_CWK_100-202425.xlsx” in the
current working directory. Then, we can use the `enter_bls_grades`
command from `blsmarks` package as follows to enter these grades into
the spreadsheet:

``` r
library(blsmarks)
enter_bls_grades(grades_df, 
                 "PSYC40545_29120_M01_CWK_100-202425.xlsx", 
                 id_col = 'nid', grade_col = 'grade')
```

This writes the grades in `grades_df` to the Excel file. Note how we
indicate which columns of the data frame have the student IDs and the
grades.

A few other things to note:

- Sometimes grades already exist in the spreadsheet that we get from BLS
  student records. These grades should not be altered, and so this
  command does not do so. It could happen that a grade exists already
  for a student for whom you have a new grade in `grades_df`. This could
  happen because the student already did and passed the element or
  module but mistakenly did the assessment again. If grades already
  exist, they are not overwritten. However, the command will print out
  the names and the old and new grades of these students for your
  information.
- Sometimes grades exist in the grades data frame but there is no row
  for the relevant student or students in the BLS marksheet. In this
  case, thise new students are *not* addded to the spreadsheet, but
  again, the command will print out the names and grades of these
  students for your information.
- Apart from this printed out information, the command will print out
  the number of students in originally in the BLS spreadsheet, how many
  of these students already have grades, how many grades are in the
  grades data frame, and how many of the grades in that data frame it
  entered into the spreadsheet.

Also note that the command `enter_bls_grades` will not overwrite an
existing grade in the spreadsheet.

## Work in progress; Use with caution

It should go without saying that this is a work in progress and so users
should be careful.

It is wise to do two things:

1.  Make a backup of the BLS spreadsheet before running the above
    command. That command will write content to the spreadsheet. It
    might do something awry, so having a backed up copy of the original
    spreadsheet means that in the worst case scenario you can just
    delete the messed up spreadsheet and go back to the original.
2.  Carefully check whether it worked and that it did what you wanted it
    or expected it to do.
