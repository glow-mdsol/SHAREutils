SHAREutils
============

This is a selection of tools for slicing and dicing the information in the CDISC Share Content Spreadsheets

* assign\_c\_codes.py - using one Spreadsheet as a reference, copy the assigned C-Codes to the cmd line in the running order for a second spreadsheet.  This can then be pasted into the content spreadsheet.

* checkheaders.py - dump out the headers from the content spreadsheet templates.

* variablemapper.py - extract variable names, c-codes and definitions out from a selection of content template sheets into a spreadsheet, so NCI representative can update in one place.

* check\_content\_sheet.py - checks content of Content Template sheets against some rules defined by the team

TODO
----
* Populate the Content Template sheets with codes inplace
* turn this into a portal/SOA type arrangement (ie POST content template sheet, get back corrected version)
* Add Sharepoint integration (to automatically download and process content templates)

