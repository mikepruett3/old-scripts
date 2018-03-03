Dim xlApp, xlBook, xlSht 
Dim filename, value1, value2, value3, value4

filename = "c:\test.xls"

Set xlApp = CreateObject("Excel.Application")
set xlBook = xlApp.WorkBooks.Open(filename)
set xlSht = xlApp.activesheet

value1 = xlSht.Cells(2, 1)
value2 = xlSht.Cells(2, 2)

'the MsgBox line below would be commented out in a real application
'this is just here to show how it works...
msgbox "Values are: " & value1 & ", " & value2 

xlBook.Close False
xlApp.Quit

'always deallocate after use...
set xlSht = Nothing
Set xlBook = Nothing
Set xlApp = Nothing 

 