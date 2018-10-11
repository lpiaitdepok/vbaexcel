'www.bluepecantraining.com
Dim ws As Worksheet
Set ws = ActiveSheet
Dim tbl As ListObject
Set tbl =ws.ListObjects("Sales_Table")
Dim newrow As ListRow
Set newrow = tbl.ListRows.Add
With newrow
    .Range(1) = 83473
    .Range(2) = "HJU -64448"
    .Range(3) = 5
End With
