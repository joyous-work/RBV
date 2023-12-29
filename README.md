# Book First

#### This branch repo shows how to choose a random verse from the books of the New Testament (NT).

Level: Assumes basic familiarity with Microsoft Excel 2019. 

#### Contains two pdfs with steps to reproduce the two macro-enabled workbooks (.xlsm).

#### A random NT verse is picked after choosing a random chapter within a book selected at random. 

* Three calls to randbetween are made for each random verse; book -> chapter -> verse.

* Uses two different Excel methods -- Index and Vlookup -- to delimit the randbetween calls. 

* Other demonstrated functionalities include:
  * Saving to and loading from CSV file
  * Checking data integrity after transformation
  * Conditional formatting
  * Named cells and ranges
  * Developer mode: Form control (button) with associated macro
  * Dynamic construction of Hyperlink to access text from online site

#### Steps outlined in RNTV-Index.pdf:
* 1A Enter Raw Data
* 1B Transform and Check Data
* 2A Random Book
* 2B Random Chapter and Verse
* 2C Add Button

#### Steps outlined in RNTV-Vlookup.pdf:
* 1A Load Data
* 1B Prepare Data
* 2A Random Book, Chapter and Verse
* 2B Add Button

Note: The last step (add macro and button) is not strictly necessary as hitting F9 would also recalculate a new random verse.
