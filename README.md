# ExcelAlmaLookup
An Excel VBA macro for performing Alma catalog lookups

To install the “Look up in Local Catalog” plugin, simply run the installer as the user who will be using the plugin (be sure to quit Excel before doing so).  After the plugin is installed, a new tab “Library Tools” will appear in the ribbon.  This tab will contain a button labeled “Look Up in Local Catalog”.  

Before clicking the button, highlight the cells containing the values you want to search for.  You can highlight an entire column, or just specific cells.  But all the values should be contained in the same column.

After highlighting the desired cells, click the “Look Up in Local Catalog” button.  The following dialog box will appear:





Below is an explanation of the fields in this dialog:

**Base URL for Alma SRU**: For Princeton, this would be “https://princeton.alma.exlibrisgroup.com/view/sru/01PRI_INST”.

**Select a range of cells to look up**: This field indicates which cells contain the values to be searched.  If you selected a range of cells before clicking the button, then this field will already contain the appropriate value.  However, it is possible to select a new range of cells by clicking the button to the right of this field.

**Validate and search equivalent ISBN/SNs**: If checked, and if “Field to search” is set to “ISBN” or “ISSN”, then each ISBN/SN will be validated.  If invalid, the check digit will be recalculated before searching. For ISBNs, the search will be done on both the 10-digit and 13-digit forms, regardless of which form is found in the spreadsheet.

**Include suppressed records**: If checked, then suppressed records will be included in the search results.

**Leftmost result column**: The column that will be populated with the first result type.  If more than one result type is selected, the others will be put in consecutive columns to the right of the first.  By default, the first empty column to the right of the spreadsheet data is selected.  Use the arrow buttons to select a different column.  If the selected column contains data, it will be overwritten. Search results will be placed in the same rows as the corresponding search values.

**Field to search**: This indicates what kind of values are in the selected cells.  Possible options are ISBN, ISSN, Call Number, Title, and MMS ID.  If an ISBN search is done, then spaces, dashes and parenthetical comments (e.g. “(paperback)”) are removed from the value before searching.  Currently, the title search does not strip stopwords or do anything else to “clean up” the titles before searching.  Thus, title searching will not be as accurate as the other search types.

**Result type(s)**:  The type of data to retrieve from the records that are found.  Selecting “True/False” will populate the column with TRUE and FALSE values based on whether the search values were found in the catalog (This kind of lookup runs much faster than the others).  The menu includes other result types, such as call numbers and location codes.  Besides the options in the menu, you can also retrieve any field from the retrieved bib record.  To retrieve an entire MARC field, enter its 3-digit tag number (e.g. “245”).  A subfield can be retrieved by appending “$” followed by the subfield code (e.g. “245$a”).  To retrieve the part of an 880 field corresponding to another field or subfield, append “-880” (e.g. “245-880” or “245-880$a”).  Subfield tags will be removed from the result before writing it to the spreadsheet.  Multiple result types can be selected, in which case they will be placed in consecutive columns in the spreadsheet, starting with the one indicated in the “Leftmost result column” field.  Use the “Add”, “Remove”, “Move Up” and “Move Down” button to edit or reorder the result types.

After specifying these values, click “OK” to begin the lookup process.  You will see the plugin populating the result column with the retrieved values.  If a record contains multiple instances of the desired result field/subfield (or, if a call number/location search is done and a record has multiple holdings), then all instances will be placed in the result cell, separated by “broken vertical bar” characters (¦).  If multiple bib records are retrieved by a single search value, the desired field from each record will be placed in the result cell, separated by solid vertical bars (|). 

