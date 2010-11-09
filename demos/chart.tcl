# $Id$
#
# This example controls Excel.  It performs the following steps.
#       - Start Excel application.
#       - Create a new workbook.
#       - Put values into some cells.
#       - Create a chart.

package require tcom

set application [::tcom::ref createobject "Excel.Application"]
$application Visible 1

set workbooks [$application Workbooks]
set workbook [$workbooks Add]
set worksheets [$workbook Worksheets]
set worksheet [$worksheets Item [expr 1]]

set data [list \
    [list "North" "South" "East" "West"] \
    [list 5.2 10.0 8.0 20.0] \
]
set sourceRange [$worksheet Range "A1" "D2"]
$sourceRange Value $data

set charts [$workbook Charts]
set chart [$charts Add]
$chart ChartWizard $sourceRange 5 [::tcom::na] 1 1 0 0 "Sales Percentages"

# Prevent Excel from prompting to save the document on close.
$workbook Saved 1
