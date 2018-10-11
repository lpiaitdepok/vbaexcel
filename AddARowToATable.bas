Dim ws As Worksheet
Set ws = ActiveSheet
Dim tbl As ListObject
Set tbl = ws.ListObjects("Sales_Table")
‘add a row at the end of the table
tbl.ListRows.Add
‘add a row as the fifth row of the table (counts the headers as a row)
tbl.ListRows.Add 5
