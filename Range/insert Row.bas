'powerspreadsheets.com
Sub insertRow()
'Range.EntireRow
'expression.Offset(RowOffset, ColumnOffset)
'Worksheet.Cells(RowIndex, ColumnIndex).EntireRow.Insert Shift:=xlShiftDown  CopyOrigin:=xlInsertFormatOriginConstant
'
'xlShiftDown or -4121: Shifts cells down.
'xlShiftToRight or -4161: Shifts cells to the right.
 
Worksheets("Insert row").Rows(6).Insert Shift:=xlShiftDown

End Sub

Sub insertMultipleRows()
'    lastRow# – firstRow# + 1
'In this example:
'    15 – 11 + 1 = 5
Worksheets(“Insert row").Rows(“11:15").Insert Shift:=xlShiftDown
End Sub

Sub insertRowFormatFromAbove()
'    xlFormatFromLeftOrAbove or 0: Newly-inserted cells take formatting from cells above or to the left.
'    xlFormatFromRightOrBelow or 1: Newly-inserted cells take formatting from cells below or to the right.

Worksheets("Insert row").Rows(21).Insert Shift:=xlShiftDown, CopyOrigin:=xlFormatFromLeftOrAbove

End Sub

Sub insertRowWithoutFormat()
'Range.ClearFormats

Dim myNewRowNumber As Long
myNewRowNumber = 31

With Worksheets("Insert row")

.Rows(myNewRowNumber).Insert Shift:=xlShiftDown

.Rows(myNewRowNumber).ClearFormats

End With

End Sub

Sub insertRowBelowActiveCell()

ActiveCell.Offset(1).EntireRow.Insert Shift:=xlShiftDown

End Sub

Sub insertCopiedRow()

With Worksheets(“Insert row")

.Rows(45).Copy
.Rows(41).Insert Shift:=xlShiftDown

End With

Application.CutCopyMode = False

End Sub

Sub insertBlankRowsBetweenRows()

Dim myFirstRow As Long
Dim myLastRow As Long
Dim myWorksheet As Worksheet
Dim iCounter As Long

myFirstRow = 5

Set myWorksheet = Worksheets(“Insert blank rows")

myLastRow = myWorksheet.Cells.Find( _
What:="*", _
LookIn:=xlFormulas, _
LookAt:=xlPart, _
SearchOrder:=xlByRows, _
SearchDirection:=xlPrevious).Row

For iCounter = myLastRow To (myFirstRow + 1) Step -1

myWorksheet.Rows(iCounter).Insert Shift:=xlShiftDown

Next iCounter

End Sub
