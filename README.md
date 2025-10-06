# profcheck-xlsx
This script converts text report from ArgyllCMS:profcheck to Excel workbook for IT8.7/2 target for better visual representation. This workbook contains 3 worksheets.

First worksheet is with IT8.7/2 target layout. Each cell/patch in it is filled with its reference color and assigned its dE value from the report with text being colored by the corresponding grade from `dE_gradations` variable.

Second worksheet is also with IT8.7/2 target layout. But each color patch is split vertically with reference color being on the left side and measured color being on the right side. Grayscale patches are split horizontally with reference color being on top and measured color being on bottom.

Third worksheet contains all patches distributed to corresponding groups based on their dE value.

Color fill and color grading can be disabled by corresponding arguments. See argparser's help.
