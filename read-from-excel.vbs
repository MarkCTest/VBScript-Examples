' This is a simple example of how toread from an excel file
' You will need to have an Excel file named 'test.xls' and a tab in there named 'TestData'

Option Explicit

Dim xlApp
Dim xlBook
Dim xlSheet

Set xlApp = CreateObject("Excel.Application")
xlApp.visible = false
Set xlBook = xlApp.Workbooks.open("\\file\location\on\your\system\test.xls")
Set xlSheet = xlBook.Worksheets("TestData")

MsgBox xlSheet.Cells(1, 1).Value
