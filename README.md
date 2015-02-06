Set of Scripts to Download Dept. of Revenue Data
================================================

##Overview
I wanted to practice getting data in R, so I created this combination of R and Excel
VBA code to download all the data and aggregate it into one Excel sheet.  The data
to download comes from the North Carolina Department of Revenue data site.

http://www.dornc.com/publications/monthlysales.html

They're nice enough to have consistent file names, and moderately consistent file
structures. Here's how the R code and Excel Macro work together.

##R Code
I used R here to download the files because I think it's easier than downloading
the files using Excel.  The only things you have to change in the R file is the 
years vector, in case you don't want them all.  The DOR data is available in a nice
Excel file from July 2000 to November 2014, as of this writing.  If a date you have 
in the loop doesn't exist, it will put a message in the Results loop.
NOTE: July 2003 was amended and given a non-standard file name.

##Excel File
Once the data is downloaded, the Excel VBA will open each file, copy the contents,
and paste them into one file.  This more or less works, with a few issues.  Those
issues are documented in the Repo.