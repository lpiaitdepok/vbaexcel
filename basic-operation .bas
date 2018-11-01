'analysistabs.com

'#Range and Format
'It will use the A5 value in the message Box
Range("A5")

'You can also use Cell Object to refer A5 as shown below:
Cells(5, 1) 'Here 5 is Row number and 1 is Column number

'It will enter the data into B5
Range("B5") = "Hello World! using Range"

'You can also use Cell Object as shown below B6:
Cells(6, 2) = "Hello World! using Cell" 'Here 6 is Row number and 2 is Column number

'You can set the background color using Interior.ColorIndex Property
Range("B1:B5").Interior.ColorIndex = 5 ' 5=Blue

'You can set the background color using Interior.ColorIndex Property
Range("B1:B10").Font.ColorIndex = 3 ' 3=Red

'You can use UCase function to change the text into Upper Case
Range("C2").Value = UCase(Range("C2").Value)

 'You can use LCase function to change the text into Upper Case
Range("C3").Value = LCase(Range("C3").Value)

'You can use Copy method
Range("A1:D11").Copy Destination:=Range("H1")

'You can use Hidden Propery of Rows
Rows("12:15").Hidden = True 'It will Hide the Rows 12 to 15

Rows("12:15").Hidden = False 'It will UnHide the Rows 12 to 15

'You can use Hidden Propery of Columns
Columns("E:G").Hidden = True 'It will Hide the Rows E to G

Columns("E:G").Hidden = False 'It will UnHide the Rows E to G

'You can use Insert and Delete Properties of Rows
Rows(6).Insert 'It will insert a row at 6 row

Rows(6).Delete 'it will delete the row 6

'You can use Insert and Delete Properties of Columns
Columns("E").Insert 'it will insert the column at E

Columns("E").Delete 'it will delete the column E

'You can use Hidden Propery of Rows
Rows(12).RowHeight = 33

Columns(5).ColumnWidth = 35

'You can use Merge Property of range
Range("E1:E5").Merge

'You can use UnMerge Property of range
Range("E1:E5").UnMerge

'# Sheet and Workbook
'You can use Select Method to select
Sheet2.Select

'You can use Acctivate Method to activate
Sheet1.Activate

'You can use ActiveSheet.Name property to get the Active Sheet name
MsgBox ActiveSheet.Name

'You can use ActiveWorkbook.Name property to get the Active Workbook name
MsgBox ActiveWorkbook.Name

'You can use Add method of a worksheet
Sheets.Add

'You can use Name property of worksheet
ActiveSheet.Name = "Temporary Sheet"

'You can use Delete method of a worksheet
Sheets("Temporary Sheet").Delete

'You can use Add method of a Workbooks
Workbooks.Add

'You can use refer parent and child object to access the range
ActiveWorkbook.Sheets("Sheet1").Range("A1") = "Sample Data"

'It will save in the deafult folder, you can mention the full path as "c:\Temp\MyNewWorkbook.xls"
ActiveWorkbook.SaveAs "MyNewWorkbook.xls"

ActiveWorkbook.Close
