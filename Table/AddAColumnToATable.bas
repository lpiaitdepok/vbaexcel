'Add a Column to a Table
'www.bluepecantraining.com

Dim ws As Worksheet
Set ws = ActiveSheet
Dim tbl As ListObject
Set tbl = ws.ListObjects("Sales_Table")
'add a new column as the 5th column in the table
tbl.ListColumns.Add(5).Name = "TAX"
'add a new column at the end of the table
tbl.ListColumns.Add.Name = "STATUS"
