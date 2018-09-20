Sub Delete_Rows(Data_range As Range, Text As String)
' case insensitive
Dim Row_Counter As Integer
For Row_Counter = Data_range.Rows.Count To 1 Step -1
If Data_range Is Nothing Then
  Exit Sub
End If
If UCase(Left(Data_range.Cells(Row_Counter, 1).Value, Len(Text))) = UCase(Text) Then
     Data_range.Cells(Row_Counter, 1).EntireRow.Delete
End If
Next Row_Counter
 
End Sub
