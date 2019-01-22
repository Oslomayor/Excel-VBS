# Excel-VBS
用VBS脚本自动处理Excel业务

## Hellow World Demo

```VBS
Dim app,workbook,sheet1

' Creat the Excel Object
Set app = WScript.CreateObject("Excel.Application")
app.Visible = True
' Creat a workbook with sheet1
Set workbook = app.Workbooks.Add
Set sheet1 = workbook.Worksheets(1)

sheet1.Range("A1").Value = "Hello World!"
```
Windows 系统内嵌VBS解释器，不需安装配置环境。将以上代码保存为后缀名为*.vbs的文本文件，双击直接可以运行。  
![](https://raw.githubusercontent.com/Oslomayor/Markdown-Imglib/master/Imgs/excel-helloworld.png)  

## Fill numbers  
see file  excel-fillnumber.vbs  
![](https://raw.githubusercontent.com/Oslomayor/Markdown-Imglib/master/Imgs/excel-fillnumber.png)  
