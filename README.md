# FIT Outstanding
A visual basic macro for Excel to read CSV files and remove lines for matched sample IDs. For best (intended) user experience, use a barcode scanner, but typing sample IDs also works.

## Structure
An Excel workbook containing one worksheet: Main, and an embedded VBA macro. 

Main contains: 
* A large, merged cell, that serves as an input field and calls a macro routine on cell change (Return / Delete),
* A form button titled "Reload" begins the macro's routine to gather data from the CSV files,
* Output fields, updated by the macro.

Main also has an embedded macro for importing and manipulating CSV data.

## Setup
Edit Sheet1 macro in VBAProject:
The global Const defaultPath should be updated from "<YOUR PATH HERE>" to an appropriate, absolute path, including trailing slashes, as exampled by the comments. 
The global Consts defaultFITFile and defaultTURFITFile should be changed to whatever the files will be named by defualt, including any extensions the user or LIMS might append. An extensionless file is perfectly acceptable, so long as the contents match the intended format, described below. 
Save as an xlsm file or xlst to prevent users saving over it inadvertently. 

## Usage
Using the LIMS outstanding gather function, print the FIT and TURFIT oustanding lists to the location and filenames specified in standard operating procedures. (These will match the defaults set in setup).
Open the macro workbook directly from Windows File Explorer. (Opening within Excel's Open File dialog enables users to edit and save the file which may impact the macro). Enable macros if asked to do so.
Click the Reload button to load the outstanding lists.
Type (or scan) a sample ID of a sample physically in your posession, into the yellow cell. Samples not found in either list will be show on the Main worksheet - these need investigating as "not booked in" or already completed and not filed away.
Select the FIT or TURFIT worksheets. Samples found will be greyed out, leaving samples that have yet to be scanned, visible - these will be samples booked in but not in your posession and should be investigated as "lost in lab" or "wrong number on tube." 

## Notes for code reviewers
Thanks for taking the time to look at this very niche macro. 
I'm uploading this for preservation, to see how a public github repository looks, and as a tiny example of my VBA experience; I appreciate neither its scale, nor the content is portfolio worthy. 
This macro was originally written for a single output file but modified for two separate files with slightly different formats. The gather was then unified (but still output to two files), so some code remains where ifs change column or row references.
Some of the code, such as CheckData could be refactored to reduce the repetition of processing lastFit then lastTurfit. Given the basic functionality, low-processing demand, and fact the macro works, no further work has been considered necessary.
