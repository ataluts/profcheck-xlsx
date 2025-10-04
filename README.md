# profcheck-xlsx
This script converts text report from ArgyllCMS:profcheck to Excel workbook for IT8.7/2 target for better visual representation. In this workbook there are 2 worksheets.

First worksheet is with IT8.7/2 target layout. Each cell/patch in it is filled with its reference color and assigned its dE value from the report with text being colored by the corresponding grade from `dE_gradations` variable.

In the second worksheet all patches are distributed to corresponding groups based on their dE value.

Color fill and color grading can be disabled by corresponding arguments. See argparser's help.
