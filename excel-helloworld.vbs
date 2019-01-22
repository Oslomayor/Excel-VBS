' https://github.com/Oslomayor/Excel-VBS
' Jan 22th, 2019 @ ICRD
Option Explicit

Dim app,workbook,sheet1

' Creat the Excel Object
Set app = WScript.CreateObject("Excel.Application")
app.Visible = True
' Creat a workbook with sheet1
Set workbook = app.Workbooks.Add
Set sheet1 = workbook.Worksheets(1)

sheet1.Range("A1").Value = "Hello World!"
sheet1.Range("A2").Value = "You are awesome!"
sheet1.Range("A3").Value = "Ù¯ÀÏ½á¹÷£¡"