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
* Load demo data
*-------------------------------------------------------------------------------
sysuse auto

*-------------------------------------------------------------------------------
* Make a copy of the pre-formatted Excel template (so that we leave the
* template unchanged)
*-------------------------------------------------------------------------------


*-------------------------------------------------------------------------------
* Write data to our COPY of the Excel template (so that we leave the original
* template unchanged)
*-------------------------------------------------------------------------------
