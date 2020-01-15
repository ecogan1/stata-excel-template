*-------------------------------------------------------------------------------
* Export data to a pre-formatted Excel file with conditional formatting
*-------------------------------------------------------------------------------

*-------------------------------------------------------------------------------
* Program Setup
*-------------------------------------------------------------------------------
version 15                // Set version number for backward compatibility
clear all                 // Start with a clean slate
macro drop _all           // Clear all macros (careful: if the do-file accepts
                          // arguments, they will be removed)
set type double           // Use double type. It's more precise than float.
set more off              // Disable partitioned output
capture log close master  // Close any opened log files
set linesize 80           // Set logfile line length
log using master.log, text replace name(master) // Begin logging
*-------------------------------------------------------------------------------

*-------------------------------------------------------------------------------
* Set up some demo data and fill it with colors (green, yellow, ...) and trend
* arrows.
*-------------------------------------------------------------------------------
// Load Stata default demo data about cars
sysuse auto

// Fill data with some rating colors
label define ratings 1 "green" 2 "yellow" 3 "orange" 4 "red" 5 "gray"
label values rep78 ratings

// Fill data with some trend arrows
label define trends 1 "↑" 2 "➚" 3 "→" 4 "↓"
recast int headroom, force
label values headroom trends

*-------------------------------------------------------------------------------
* Make a copy of the pre-formatted Excel template (so that we leave the
* original template intact)
*-------------------------------------------------------------------------------
copy template.xlsx sdgindex.xlsx, replace

*-------------------------------------------------------------------------------
* Write data to our copy of the template.
* We must use sheetmodify in combination with keepcellfmt. Otherwise, we would
* only retain conditional formatting and "normal", unconditional formatting
* (such as bold, underline, angled) would be lost.
*-------------------------------------------------------------------------------
export excel sdgindex.xlsx,           /// write to our copy of the template
             sheetmodify              /// keep all formatting
             keepcellfmt              /// keep all formatting
             firstrow(varlabels)      /// write variable labels
             sheet("Overall Results") //  the sheet to write to

// Write data to another sheet of our copy of the template. This is just for
// demo purposes to show that formatting is indeed retained for the various
// sheets.
export excel sdgindex.xlsx,               /// write to our copy of the template
             sheetmodify                      /// keep all formatting
             keepcellfmt                      /// keep all formatting
             firstrow(varlabels)              /// write variable labels
             sheet("Values, Ratings, Trends") //  the sheet to write to
