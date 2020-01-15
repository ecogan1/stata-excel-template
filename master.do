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

// Clone the variables to test that formatting still works in the later
// columns as well
clonevar more_ratings = rep78
clonevar more_trends = headroom

*-------------------------------------------------------------------------------
* Make a copy of the pre-formatted Excel template (so that we leave the
* original template intact)
*-------------------------------------------------------------------------------
copy template.xlsx sdgindex.xlsx, replace

*-------------------------------------------------------------------------------
* Write data to our copy of the template.
* Unfortunately, we can only retain the conditional formatting applied to the
* template. "Normal" formatting is lost.
* There is a way to keep "normal" formatting for cells that have text in the
* template, by using the sheetmodify and keepcellfmt options. However, that is
* not super useful, so we instead use the sheetreplace option here and then
* use Stata's excelput to bold and slant the header row.
*-------------------------------------------------------------------------------
export excel sdgindex.xlsx,           /// write to our copy of the template
             sheetreplace             /// keep only conditional formatting
             firstrow(varlabels)      /// write variable labels
             sheet("Overall Results") //  the sheet to write to

// Format the header row
putexcel set sdgindex.xlsx, modify sheet("Overall Results")
putexcel (A1:ZZ1),                    /// target the entire first row
         overwritefmt                 /// remove any existing formatting
         bold                         /// bold the header row
         txtrotate(45)                //  angle/slant the text
putexcel clear

// Write data to another sheet of our copy of the template. This is just for
// demo purposes to show that formatting is indeed retained for the various
// sheets.
export excel sdgindex.xlsx,           /// write to our copy of the template
             sheetreplace             /// keep only conditional formatting
             firstrow(varlabels)      /// write variable labels
             sheet("Values, Ratings, Trends") //  the sheet to write to

// Format the header row
putexcel set sdgindex.xlsx, modify sheet("Values, Ratings, Trends")
putexcel (A1:ZZ1),                    /// target the entire first row
         overwritefmt                 /// remove any existing formatting
         bold                         /// bold the header row
         txtrotate(45)                // angle/slant the text
putexcel clear
