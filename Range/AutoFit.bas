'This example changes the width of columns A through I on Sheet1 to achieve the best fit.
Worksheets("Sheet1").Columns("A:I").AutoFit

'This example changes the width of columns A through E on Sheet1 to achieve the best fit, based only on the contents of cells A1:E1.
Worksheets("Sheet1").Range("A1:E1").Columns.AutoFit

'https://docs.microsoft.com/
