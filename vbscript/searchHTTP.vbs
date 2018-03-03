Dim http,ReadCacheFile,fso,NewURL,TextStream
Const ForReading = 1, ForWriting = 2, ForAppending = 3
Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0

set fso=CreateObject("Scripting.FileSystemObject")
'set ReadCacheFile=fso.OpenTextFile("c:\test.txt",TRUE)
ReadCacheFile = fso.FileExists("C:\test.txt")

set TextStream = ReadCacheFile.OpenAsTextStream(ForReading,TristateUseDefault)

Do While Not TextStream.AtEndOfStream
   Dim Line
   Line = TextStream.readline
   WScript.Echo Line
Loop

ReadCacheFile.Close

