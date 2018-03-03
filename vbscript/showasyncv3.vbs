servername = wscript.arguments(0)
set shell = createobject("wscript.shell")
strValueName = "HKLM\SYSTEM\CurrentControlSet\Control\TimeZoneInformation\ActiveTimeBias"
minTimeOffset = shell.regread(strValueName)
toffset = datediff("h",DateAdd("n", minTimeOffset, now()),now())

report = "<table border=""1"" width=""100%"">" & vbcrlf
report = report & "  <tr>" & vbcrlf
report = report & "<td align=""center"" bgcolor=""#000080""><b><font color=""#FFFFFF"">DisplayName</font></b></td>" & vbcrlf
report = report & "<td align=""center"" bgcolor=""#000080""><b><font color=""#FFFFFF"">Email Address</font></b></td>" & vbcrlf
report = report & "<td align=""center"" bgcolor=""#000080""><b><font color=""#FFFFFF"">Device Type</font></b></td>" & vbcrlf
report = report & "<td align=""center"" bgcolor=""#000080""><b><font color=""#FFFFFF"">Device ID</font></b></td>" & vbcrlf
report = report & "<td align=""center"" bgcolor=""#000080""><b><font color=""#FFFFFF"">FolderSync</font></b></td>" & vbcrlf
report = report & "<td align=""center"" bgcolor=""#000080""><b><font color=""#FFFFFF"">ContactSync</font></b></td>" & vbcrlf
report = report & "<td align=""center"" bgcolor=""#000080""><b><font color=""#FFFFFF"">CalendarSync</font></b></td>" & vbcrlf
report = report & "<td align=""center"" bgcolor=""#000080""><b><font color=""#FFFFFF"">autdstate.xml</font></b></td>" & vbcrlf
report = report & "</tr>" & vbcrlf
set req = createobject("microsoft.xmlhttp")
set com = createobject("ADODB.Command")
set conn = createobject("ADODB.Connection")
Set iAdRootDSE = GetObject("LDAP://RootDSE")
strNameingContext = iAdRootDSE.Get("configurationNamingContext")
strDefaultNamingContext = iAdRootDSE.Get("defaultNamingContext")
Conn.Provider = "ADsDSOObject"
Conn.Open "ADs Provider"
polQuery = "<LDAP://" & strNameingContext &  ">;(&(objectCategory=msExchRecipientPolicy)(cn=Default Policy));distinguishedName,gatewayProxy;subtree"
svcQuery = "<LDAP://" & strNameingContext & ">;(&(objectCategory=msExchExchangeServer)(cn=" & Servername & "));cn,name,legacyExchangeDN;subtree"
Com.ActiveConnection = Conn
Com.CommandText = polQuery
Set plRs = Com.Execute
while not plRs.eof
	for each adrobj in plrs.fields("gatewayProxy").value
		if instr(adrobj,"SMTP:") then dpDefaultpolicy = right(adrobj,(len(adrobj)-instr(adrobj,"@")))
	next
	plrs.movenext
wend
wscript.echo dpDefaultpolicy 
Com.CommandText = svcQuery
Set Rs = Com.Execute
while not rs.eof	
	GALQueryFilter =  "(&(&(&(& (mailnickname=*)(!msExchHideFromAddressLists=TRUE)(| (&(objectCategory=person)(objectClass=user)(msExchHomeServerName=" & rs.fields("legacyExchangeDN") & ")) )))))"
	strQuery = "<LDAP://"  & strDefaultNamingContext & ">;" & GALQueryFilter & ";displayname,mail,distinguishedName,mailnickname,proxyaddresses;subtree"
	com.Properties("Page Size") = 100
	Com.CommandText = strQuery
	Set Rs1 = Com.Execute
	while not Rs1.eof
		falias = "https://" & servername & "/exadmin/admin/" & dpDefaultpolicy & "/mbx/"
		if not isnull(rs1.fields("proxyaddresses").value) then 
			for each paddress in rs1.fields("proxyaddresses").value
				if instr(paddress,"SMTP:") then falias = falias & replace(paddress,"SMTP:","")  & "/non_ipm_subtree"
			next
			wscript.echo  falias 
			SerachAsync(falias)
		else 
			wscript.echo "*** Null Proxy ****	: " & rs1.fields("mailnickname")
		end if
		rs1.movenext
	wend
	rs.movenext
wend
rs.close
set conn = nothing
set com = nothing
report = report & "</table>" & vbcrlf
Set fso = CreateObject("Scripting.FileSystemObject")
set wfile = fso.opentextfile("c:\temp\asreport.htm",2,true) 
wfile.write report
wfile.close
set wfile = nothing
set fso = nothing

wscript.echo "Done"

sub SerachAsync(furl)
strQuery = "<?xml version=""1.0""?><D:searchrequest xmlns:D = ""DAV:"" >"
strQuery = strQuery & "<D:sql>SELECT ""http://schemas.microsoft.com/mapi/proptag/x3001001E"""
strQuery = strQuery & " FROM scope('shallow traversal of """
strQuery = strQuery & furl & """') Where ""DAV:ishidden"" = False AND ""DAV:isfolder"" = True AND "
strQuery = strQuery & """http://schemas.microsoft.com/mapi/proptag/x3001001E"" = 'Microsoft-Server-ActiveSync'</D:sql></D:searchrequest>"
req.open "SEARCH", furl, false
req.setrequestheader "Content-Type", "text/xml"
req.setRequestHeader "Translate","f"
on error resume next
req.send strQuery
if err.number <> 0 then wscript.echo err.description
on error goto 0
If req.status >= 500 Then
ElseIf req.status = 207 Then
	set oResponseDoc = req.responseXML
	set oNodeList = oResponseDoc.getElementsByTagName("d:x3001001E")
	if oNodeList.length <> 0 then
		wscript.echo "Active-Sync Folder Exists"
		displayAyncSub(furl & "/Microsoft-Server-ActiveSync")
	else
		wscript.echo "No Active-Sync Folder"
	end if 
Else
End If
	
end sub

sub displayAyncSub(furl)

strQuery = "<?xml version=""1.0""?><D:searchrequest xmlns:D = ""DAV:"" >"
strQuery = strQuery & "<D:sql>SELECT ""http://schemas.microsoft.com/mapi/proptag/x3001001E"""
strQuery = strQuery & " FROM scope('shallow traversal of """
strQuery = strQuery & furl & """') Where ""DAV:ishidden"" = False AND ""DAV:isfolder"" = True</D:sql></D:searchrequest>"
req.open "SEARCH", furl, false
req.setrequestheader "Content-Type", "text/xml"
req.setRequestHeader "Translate","f"
on error resume next
req.send strQuery
if err.number <> 0 then wscript.echo err.description
on error goto 0
If req.status >= 500 Then
ElseIf req.status = 207 Then
	set oResponseDoc = req.responseXML
	set oNodeList = oResponseDoc.getElementsByTagName("d:x3001001E")
	for each node in oNodeList
		call displaydeviceSub(furl & "/" & node.text,node.text)
	next
Else
End If
end sub

sub displaydeviceSub(furl,fname)

strQuery = "<?xml version=""1.0""?><D:searchrequest xmlns:D = ""DAV:"" >"
strQuery = strQuery & "<D:sql>SELECT ""http://schemas.microsoft.com/mapi/proptag/x3001001E"""
strQuery = strQuery & " FROM scope('shallow traversal of """
strQuery = strQuery & furl & """') Where ""DAV:ishidden"" = False AND ""DAV:isfolder"" = True</D:sql></D:searchrequest>"
req.open "SEARCH", furl, false
req.setrequestheader "Content-Type", "text/xml"
req.setRequestHeader "Translate","f"
on error resume next
req.send strQuery
if err.number <> 0 then wscript.echo err.description
on error goto 0
If req.status >= 500 Then
ElseIf req.status = 207 Then
	set oResponseDoc = req.responseXML
	set oNodeList = oResponseDoc.getElementsByTagName("d:x3001001E")
	for each node in oNodeList
		report = report & "<tr>" & vbcrlf
		report = report & "<td align=""center"">" & rs1.fields("displayname") & "&nbsp;</td>" & vbcrlf
		report = report & "<td align=""center"">" & rs1.fields("mail") & "&nbsp;</td>" & vbcrlf
		report = report & "<td align=""center"">" & fname & "&nbsp;</td>" & vbcrlf
		report = report & "<td align=""center"">" & node.text  & "&nbsp;</td>" & vbcrlf
		report = report & finditems(furl & "/" & node.text)
		report = report & "</tr>" & vbcrlf
	next
Else
End If
end sub

function finditems(furl)

hascalsyc = 0
hasfolsyc = 0
hasconsyc = 0
hasautd = 0
rback = ""
strQuery = "<?xml version=""1.0""?><D:searchrequest xmlns:D = ""DAV:"" >"
strQuery = strQuery & "<D:sql>SELECT ""DAV:displayname"", ""DAV:getlastmodified"""
strQuery = strQuery & " FROM scope('shallow traversal of """
strQuery = strQuery & furl & """') Where ""DAV:isfolder"" = False</D:sql></D:searchrequest>"
req.open "SEARCH", furl, false
req.setrequestheader "Content-Type", "text/xml"
req.setRequestHeader "Translate","f"
on error resume next
req.send strQuery
if err.number <> 0 then wscript.echo err.description
on error goto 0
rem wscript.echo req.responsetext
If req.status >= 500 Then
ElseIf req.status = 207 Then
	set oResponseDoc = req.responseXML
	set oNodeList = oResponseDoc.getElementsByTagName("a:displayname")
	set oNodemodlist = oResponseDoc.getElementsByTagName("a:getlastmodified")
	wscript.echo oNodeList.length
	for i = 1 to oNodeList.length
		set onode = oNodeList.nextNode
		set onode1 = oNodemodlist.nextNode
		select case lcase(onode.text)
			case "calendarsyncfile" hascalsyc = 1
						hascalsycval = DateAdd("h",toffset,(left(replace(replace(onode1.text,"T"," "),"Z",""),19)))
			case "foldersyncfile"	hasfolsyc = 1
						hasfolsycval = DateAdd("h",toffset,(left(replace(replace(onode1.text,"T"," "),"Z",""),19)))
			case "contactssyncfile" hasconsyc = 1
						hasconsycval = DateAdd("h",toffset,(left(replace(replace(onode1.text,"T"," "),"Z",""),19)))
			case "autdstate.xml"    hasautd = 1
						hasautdval = DateAdd("h",toffset,(left(replace(replace(onode1.text,"T"," "),"Z",""),19)))
		end select
	next
Else
End If
wscript.echo hasfolsyc
if hasfolsyc = 1  then
	rback = rback & "<td align=""center"">" & hasfolsycval & "&nbsp;</td>" & vbcrlf
else
	rback = rback & "<td align=""center"">No&nbsp;</td>" & vbcrlf
end if
if hasconsyc  = 1  then
	rback = rback & "<td align=""center"">" & hasconsycval & "&nbsp;</td>" & vbcrlf
else
	rback = rback & "<td align=""center"">No&nbsp;</td>" & vbcrlf
end if
if hascalsyc  <> 0  then
	rback = rback & "<td align=""center"">" & hascalsycval & "&nbsp;</td>" & vbcrlf
else
	rback = rback & "<td align=""center"">No&nbsp;</td>" & vbcrlf
end if
if hasautd  <> 0  then
	rback = rback & "<td align=""center"">" & hasautdval & "&nbsp;</td>" & vbcrlf
else
	rback = rback & "<td align=""center"">No&nbsp;</td>" & vbcrlf
end if
finditems = rback
end function


