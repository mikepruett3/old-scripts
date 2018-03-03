Dim http,NewCacheFile,fso,NewURL
NewURL=WScript.arguments.item(0)
set fso=CreateObject("Scripting.FileSystemObject")
set NewCacheFile=fso.CreateTextFile("c:\test.txt",TRUE)
set http=CreateObject("Microsoft.XMLHTTP")
http.open "GET",NewURL & "/index.asp",FALSE
http.send ""
NewCacheFile.Write http.responseText
NewCacheFile.Close
set http=nothing