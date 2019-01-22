Option Explicit

Dim app,workbook,sheet1,sheet2,sheet3
DIm row,col

' Creat the Excel Object
Set app = WScript.CreateObject("Excel.Application")
' Make it Visible
app.Visible = True
' Creat a workbook
Set workbook = app.Workbooks.Add

' Quote the sheet1
Set sheet1 = workbook.Worksheets(1)
For row = 1 To 10
    For col = 1 To 10
        ' Fill Cells with numbers 
        sheet1.Cells(row,col).Value = CInt(Int((100*Rnd())+1))
    Next
Next

' Add sheet2
Set sheet2 = workbook.Worksheets.Add
' Fill Range with Formula
sheet2.Range("A1:J10").Formula = "=Int(Rand()*100+1)"

' Add sheet3
Set sheet3 = workbook.Worksheets.Add
' Fill a specific Cell with String
sheet3.Cells(5,5).Value = "555"