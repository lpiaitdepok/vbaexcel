Sub mergeCells()
'Worksheet.Range("FirstCell:LastCell").Merge

    Worksheets("Merge Cells").Range("A5:E6").Merge

End Sub

Sub unmergeCells()
'Worksheet.Range("A1CellReference").UnMerge

    Worksheets("Merge Cells").Range("C6").UnMerge

End Sub

Sub mergeCellsAndCenter()
'HorizontalAlignment = xlCenter

    With Worksheets("Merge Cells").Range("A8:E9")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Merge
    End With

End Sub

Sub mergeCellsAcross()

    Worksheets("Merge Cells").Range("A11:E15").Merge Across:=True

End Sub

Sub mergeCellsBasedOnCriteria()

    Dim myFirstRow As Long
    Dim myLastRow As Long
    Dim myCriteriaColumn As Long
    Dim myFirstColumn As Long
    Dim myLastColumn As Long
    Dim myWorksheet As Worksheet
    Dim myCriteria As String
    Dim iCounter As Long

    myFirstRow = 5
    myCriteriaColumn = 1
    myFirstColumn = 1
    myLastColumn = 5
    myCriteria = "Merge cells"

    Set myWorksheet = Worksheets("Merge Cells Based on Criteria")

    With myWorksheet

        myLastRow = .Cells.Find(What:="*", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row

        For iCounter = myLastRow To myFirstRow Step -1
            If .Cells(iCounter, myCriteriaColumn).Value = myCriteria Then .Range(.Cells(iCounter, myFirstColumn), .Cells(iCounter, myLastColumn)).Merge
        Next iCounter

    End With

End Sub

Sub mergeCellsBasedOnCellValue()

    Dim myFirstRow As Long
    Dim myLastRow As Long
    Dim myBaseColumn As Long
    Dim mySizeColumn As Long
    Dim myWorksheet As Worksheet
    Dim iCounter As Long

    myFirstRow = 5
    myBaseColumn = 1
    mySizeColumn = 1

    Set myWorksheet = Worksheets("Merge Cells Based on Cell Value")

    With myWorksheet

        myLastRow = .Cells.Find(What:="*", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row

        For iCounter = myLastRow To myFirstRow Step -1
            .Cells(iCounter, myBaseColumn).Resize(ColumnSize:=.Cells(iCounter, mySizeColumn).Value).Merge
        Next iCounter

End With

End Sub

'powerspreadsheets.com
'Christopher Lee. Mahir Otodidak VBA Macro Excel. Elex Media Komputindo
'Flavio Morgado.Programming Excel with VBA.Apress
