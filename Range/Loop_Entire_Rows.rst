referensi:
https://support.office.com/

``visual-basic``

::

  For i = 1 To Rows.Count

  Next i

' **Note**: worksheets can have up to 65,536 rows in Excel 2003 and up to 1,048,576 rows in Excel 2007 or later. No matter what version you are using, the code line above loops through all rows (downloadable Excel file is in Excel 97-2003 format). Data in cells outside of this row and column limit is lost in Excel 97-2003.


``visual-basic``

::

  For i = 1 To Columns.Count

  Next i


' **Note**: worksheets can have up to 256 columns in Excel 2003 and up to 16,384 columns wide in Excel 2007 or later. No matter what version you are using, the code line above loops through all Columns (downloadable Excel file is in Excel 97-2003 format). Data in cells outside of this row and column limit is lost in Excel 97-2003.
