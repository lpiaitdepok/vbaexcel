'www.excel-easy.com

'Excel VBA Range Object
Range("B3").Value = 2

Range("A1:A4").Value = 5

Range("A1:A2,B3:C4").Value = 10

Range("Prices").Value = 15

Cells(3, 2).Value = 2

Range(Cells(1, 1), Cells(4, 1)).Value = 5

'Declare a Range Object
'--------------------
Dim example As Range
Set example = Range("A1:C4")

example.Value = 8
'--------------------------------------

'Select
Dim example As Range
Set example = Range("A1:C4")

example.Select
'----------------------------------------------

Worksheets(3).Activate
Worksheets(3).Range("B7").Select
'----------------------------------

Dim example As Range
Set example = Range("A1:C4")

example.Rows(3).Select
'---------------------------------

Dim example As Range
Set example = Range("A1:C4")

example.Columns(2).Select
'---------------------------------------

Range("A1:A2").Select
Selection.Copy

Range("C3").Select
ActiveSheet.Paste

Range("C3:C4").Value = Range("A1:A2").Value

'-----------------------------------------------

Range("A1").ClearContents
'or
Range("A1").Value = ""

'----------------------------------
Dim example As Range
Set example = Range("A1:C4")

MsgBox example.Count

MsgBox example.Rows.Count
MsgBox example.Columns.Count
