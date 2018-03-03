' -----------------------------------------------------------------'
' MapDrives.vbs - Map Network Drives to the File Server(s)
' Author - Mike Pruett @ http://amanojyaku.info
' Version 2.1 - October 3rd 2008
'
' Description:
'       This script is an evolution of a script was originally made
'				in DOS/Bash script. Script will backup any files from the
'				Source directory, as long as the file matches an extension in
'				the inclusion list. These files are then copied over to the
'				desired network share. NOTE!! In order for the script to work
'				This share must be accessable from the local machine. (Must
'				show up in My Computer as a	drive letter.) The script is not
'				too picky, as it does not require Block Level access to	the 
'				Drive. (NAS and SMB file sharing OK!!)
' -----------------------------------------------------------------'

'Dim DomainUsername : DomainUsername = "<Domain>\<username>"
'Dim DomainPassword : DomainPassword = "<password>"
Dim XMLFile : XMLFile = "drives.xml"
Dim oXML,oNetwork,oFSO,Nodes,Count,UserName,DiskInUse,MessageBoxText
Set oXML=CreateObject("Microsoft.XMLDOM")
Set oNetwork=CreateObject("WScript.Network")
Set oFSO=CreateObject("Scripting.FileSystemObject")
oXML.Async=False
oXML.Load(XMLFile)
UserName=oNetwork.UserName
Nodes=GetNodeCount
Do While Count <> Nodes
  Dim DriveName,DrivePath,DriveLetter,ServerName,ServerIP,DriveMapping
  GetDriveInfo Count
  PingStatus = Ping(ServerName)
  If PingStatus = False Then
    PingStatus = Ping(ServerIP)
    If PingStatus = False Then
      MessageBoxText = vbcrlf & "Could Not communicate with the File Server!!" & vbcrlf & vbcrlf & _
      "Pings to " & ServerName & "/" & ServerIP & " failed..."
      MsgBox MessageBoxText, VBCritical, "Something Went Wrong..."
      WScript.Quit
    Else
      MessageBoxText = MessageBoxText & "There were problems communicating with the File Server!" &  _
      " Using the IP Address " & ServerIP & " to establish communications..." & vbcrlf & vbcrlf
      ServerName = ServerIP
    End If
  End If
  If DrivePath="%username%" Then
    DrivePath = UserName
  End If
  DiskInUse = DriveCheck(DriveLetter)
  If DiskInUse = FALSE Then
    If DomainUsername <> "" Then
      oNetwork.MapNetworkDrive DriveLetter & ":" , "\\" & ServerName & "\" & DrivePath ,, DomainUsername , DomainPassword
    Else
      oNetwork.MapNetworkDrive DriveLetter & ":" , "\\" & ServerName & "\" & DrivePath
    End If
  Else
    MessageBoxText = MessageBoxText & "Could not map " & DriveName & " to " & DriveLetter & "... Drive Letter in use!" & vbcrlf & vbcrlf
  End If
  Count = Count + 1
Loop

If MessageBoxText <> "" Then
  MsgBox MessageBoxText, VBExclamation, "Oops, Something's Not Right..."
End If

Set oXML = Nothing
Set oNetwork = Nothing
Set oFSO = Nothing

Function GetNodeCount
  Dim Node,NodeList,NodeCount
  Set NodeList=oXML.documentElement.childNodes
  For Each Node in NodeList
    NodeCount=NodeCount + 1
  Next
  GetNodeCount = NodeCount
  Set NodeList = Nothing
  Set Node = Nothing
End Function

Sub GetDriveInfo(ItemNum)
  Dim Node,NodeList
  Set NodeList=oXML.documentElement.childNodes.item(ItemNum).childNodes
  For Each Node in NodeList
    Select Case Node.NodeName
      Case "Name"
        DriveName = Node.Text
      Case "Path"
        DrivePath = Node.Text
      Case "Letter"
        DriveLetter = Node.Text
      Case "ServerName"
        ServerName = Node.Text
      Case "IP"
        ServerIP = Node.Text
    End Select
  Next
  Set NodeList = Nothing
  Set Node = Nothing
End Sub

Function DriveCheck(TargetDrive)
  Dim Drives,DiskDrive
  Set Drives = oFSO.Drives
  For Each DiskDrive in Drives
    If DiskDrive.DriveLetter = TargetDrive Then
      DriveCheck = TRUE
    End If
  Next
  Set Drives = Nothing
  Set DiskDrive = Nothing
End Function

Function Ping(Host)
    Dim oPing,ReturnStatus
    Set oPing = GetObject("winmgmts:{impersonationLevel=impersonate}").ExecQuery("select * from Win32_PingStatus where address = '" & Host & "'")
    For Each ReturnStatus in oPing
        If IsNull(ReturnStatus.StatusCode) or ReturnStatus.StatusCode<>0 Then
            Ping = False
        Else
            Ping = True
        End If
    Next
    Set oPing = Nothing
    Set ReturnStatus = Nothing
End Function
