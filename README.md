# Verse First

#### This branch repo shows how to choose a random verse from the collection of all New Testament (NT) verses. 

Level: Assumes basic familiarity with Microsoft Excel 2019. 

#### Contains one workbook with a pdf describing steps to reproduce it.

#### A random verse is picked from the collection of all NT verses, and its BCV (Book-Chapter-Verse) coordinates are identified. 

* A single call to randbetween is made for each random verse.

* Uses two Excel functions -- PercentrankInc and Offset -- to identify the BCV coordinates of a random verse.

* Other demonstrated functionalities include:
  * Load and transform data with Power Query Editor
  * Change table style and convert table to range
  * Create running total column and rows; cumulative series
  * Copy data with special Paste options: Column widths, Values & Source Formatting
  * Construct nested functions (nesting formulas)
  * Test edge cases, and tune significance parameter of percentrank.inc
  * Construct rules and apply conditional formatting
  * Apply data validation to restrict input values to cell
  * Create dynamic ranges with offset function
  * Use named cells and ranges in formulas
  * Compare two numeric ranges for equality with subtraction
  * Display cell formulas with Formulatext
  * Limit cell editing with Locked cells and Protect sheet

#### Steps outlined in RNTV-PercentRankInc.pdf:
* 1A Load Data
* 1B Prepare Data
* 2A Book Search and Test
* 2B Chapter and Verse Search
* 2C Test Chapter and Verse Search
* 2D Add Link to Verse Text
