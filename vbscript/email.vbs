set iMsg = CreateObject("CDO.Message")
set iConf = CreateObject("CDO.Configuration")

With iConf.Fields
.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "<server address here>"
.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 10
.Update
End With

With iMsg
Set .Configuration = iConf
.To = "<To address here>"
.From = "<From Address here>"
.Subject = "<Subject Line here>"
.textBody = "<Body Here>"
.Send
End With

Set iMsg = Nothing
Set iConf = Nothing