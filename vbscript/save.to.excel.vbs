Dim xlApp, xlBook, xlSht 
Dim filename, value1, value2, value3, value4

on error resume next

filename = "c:\test.xls"

Set xlApp = CreateObject("Excel.Application")
set xlBook = xlApp.WorkBooks.Open(filename)
set xlSht = xlApp.activesheet

xlApp.DisplayAlerts = False

'write data into the spreadsheet
xlSht.Cells(4, 1) = "New Data"

xlBook.Save
xlBook.Close SaveChanges=True
xlApp.Close
xlApp.Quit

'always deallocate after use...
set xlSht = Nothing
Set xlBook = Nothing
Set xlApp = Nothing