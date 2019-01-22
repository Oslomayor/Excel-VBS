Option Explicit

Dim app,workbook,sheet1
Dim row,col,count,gap

Set app = WScript.CreateObject("Excel.Application")
app.Visible = True
Set workbook = app.Workbooks.Add
Set sheet1 = workbook.Worksheets(1)

row = 1
col = 1
gap = 0
do while row <= 30
    for count= 1 To 100
        if col <= 100 then
            sheet1.Cells(row,col).Value = count
        end if
        workbook.ActiveSheet.Cells(col).EntireColumn.AutoFit
        col = col + 1 + gap
    next
    row = row + 1
    gap = row - 1
    col = 1
loop

