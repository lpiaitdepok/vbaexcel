Sub FindingLastRow()
'Ways To Find The Last Row
'PURPOSE: Different ways to find the last row number of a range
'SOURCE: www.TheSpreadsheetGuru.com

Dim sht As Worksheet
Dim LastRow As Long

Set sht = ActiveSheet

'Using Find Function (Provided by Bob Ulmas)
  LastRow = sht.Cells.Find("*", searchorder:=xlByRows, searchdirection:=xlPrevious).Row

'Using SpecialCells Function
  LastRow = sht.Cells.SpecialCells(xlCellTypeLastCell).Row

'Ctrl + Shift + End
  LastRow = sht.Cells(sht.Rows.Count, "A").End(xlUp).Row

'Using UsedRange
  sht.UsedRange 'Refresh UsedRange
  LastRow = sht.UsedRange.Rows(sht.UsedRange.Rows.Count).Row

'Using Table Range
  LastRow = sht.ListObjects("Table1").Range.Rows.Count

'Using Named Range
  LastRow = sht.Range("MyNamedRange").Rows.Count

'Ctrl + Shift + Down (Range should be first cell in data set)
  LastRow = sht.Range("A1").CurrentRegion.Rows.Count

End Sub
