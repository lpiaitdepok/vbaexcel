Sub FindingLastColumn()
'Ways To Find The Last Column
'PURPOSE: Different ways to find the last column number of a range
'SOURCE: www.TheSpreadsheetGuru.com

Dim sht As Worksheet
Dim LastColumn As Long

Set sht = ThisWorkbook.Worksheets("Sheet1")

'Ctrl + Shift + End
  LastColumn = sht.Cells(7, sht.Columns.Count).End(xlToLeft).Column

'Using UsedRange
  sht.UsedRange 'Refresh UsedRange
  LastColumn = sht.UsedRange.Columns(sht.UsedRange.Columns.Count).Column

'Using Table Range
  LastColumn = sht.ListObjects("Table1").Range.Columns.Count

'Using Named Range
  LastColumn = sht.Range("MyNamedRange").Columns.Count

'Ctrl + Shift + Right (Range should be first cell in data set)
  LastColumn = sht.Range("A1").CurrentRegion.Columns.Count

End Sub
