'**********************************************************************************
'**********************************************************************************
'**********************************************************************************
' 
' slogon shares.vbs
'
' A vbscript to map shared folders at logon using simple commands.
' Each drive mapping invoked below records an entry in the local Application log.
' 
' Created by Rob Pennoyer Decemeber 17, 2005.
' 
' Instructions below.  
'
'**********************************************************************************
'**********************************************************************************
'**********************************************************************************


Option Explicit
Dim objGroupList, objUser, objNetwork, objShell
Dim strNetBIOSDomain, strGroup, strNTName
Dim strScriptName

Set objNetwork = CreateObject("Wscript.Network")
Set objShell = Wscript.CreateObject("Wscript.Shell")
Const EVENT_ERROR = 1, EVENT_WARNING = 2, EVENT_INFO = 4

'****** Record script start in Application event log
strScriptName = WScript.ScriptName
objShell.LogEvent EVENT_INFO, strSCriptName & " script starting." 



'**********************
'**********************
'NetBIOS Domain name -- You must specify this here:

strNetBIOSDomain = "DOMAIN"

'**********************
'**********************
'****** Loop required for Win9x clients only.  Ignore this
strNTName = ""
On Error Resume Next
Do While strNTName = ""
  strNTName = objNetwork.userName
  Err.Clear
  If Wscript.Version > 5 Then
    Wscript.Sleep 100
  End If
Loop
On Error GoTo 0


'****** Bind to the user object in Active Directory with the WinNT provider.
Set objUser = GetObject("WinNT://" & strNetBIOSDomain & "/" _
  & strNTName & ",user")


'**********************************************************************************
'**********************************************************************************
'**********************************************************************************
' Edit commands within these bars only.  
' ***Don't forget to set the NETBIOS domain name up above.
' There are three commands available for use.  Use them in any combination or
' quantity you choose.  
'
' -----> RemoveAllNetworkDrives <-----
' Disconnects all mapped network drives
' Usage:
' RemoveAllNetworkDrives
' Note: This is not a necessary step
'
' -----> MapFolder DriveLetter,Sharename <-----
' Maps a drive letter to a share
' Usage:
' MapFolder "X:","\\server\share"
' Note: there is no trailing \ charater
'
' -----> MapHomeFolder DriveLetter,HomeFolderParent <------
' Maps a folder and appends the current username.  Example below for a user named
' "bob" will map x: to \\server\users\bob
' Usage:
' MapHomeFolder "x:","\\server\users"
'
' -----> GroupMapFolder DriveLetter,ShareName,GroupName <------
' Maps a folder only if the user is a member of a specified security group
' The groupname *must* exist, even if the user is not a member, or the script will fail
' This must be a security group, not an Organizational Unit.
' Usage:
' GroupMapFolder "x:","\\server\share","ExecutiveUsers"
'


RemoveAllNetworkDrives

'MapFolder "f:","\\fileserver\shared"

'MapFolder "g:","\\otherserver\public"

'GroupMapFolder "t:","\\fileserver\accounting","AccountingGroup"

'MapHomeFolder "h:","\\otherserver\home"







'**********************************************************************************
'**********************************************************************************
'**********************************************************************************



'****** Clean up.
objShell.LogEvent EVENT_INFO, strSCriptName & " script complete." 
Set objGroupList = Nothing
Set objUser = Nothing
Set objNetwork = Nothing





Function MapHomeFolder(DriveLetter,HomeFolderParent)
' Maps a folder using the username.  

  MapFolder DriveLetter,HomeFolderParent & "\" & objNetwork.Username
End Function


Function GroupMapFolder(DriveLetter,ShareName,GroupName)
' Maps a folder only if the user is a member of a specified security group

If IsMember(GroupName) Then
  MapFolder DriveLetter,Sharename
End If
End Function


Function MapFolder(DriveLetter,ShareName)
' Maps a folder
' Usage:
' MapFolder("X:","\\server\share")

'debug
'MsgBox "MapFolder running"

Dim RunningLog, EventType
'First sleep for .2 seconds, so the commands don't fall on top of each other
  WScript.sleep 200
RunningLog = strScriptName & " is attempting to map" & vbcrlf & DriveLetter & " to " & ShareName & vbcrlf
EventType = EVENT_INFO
  On Error Resume Next
  Err.Clear
  objNetwork.MapNetworkDrive DriveLetter, ShareName
  If Err.Number <> 0 Then
    RunningLog = RunningLog & "Drive already exists!  Now attempting to remove mapping." & vbcrlf
    Err.Clear
    objNetwork.RemoveNetworkDrive DriveLetter, True, True
    if Err.Number <> 0 Then 
       RunningLog = RunningLog & "Unable to remove existing map. " & vbcrlf
       EventType = EVENT_ERROR
       Err.Clear
    Else
       RunningLog = RunningLog & "Existing drive mapping removed successfully.  Now attemping to map again." & vbcrlf
    End If
    objNetwork.MapNetworkDrive DriveLetter, ShareName
    If Err.Number <> 0 Then
       RunningLog = RunningLog & "Unable to map drive." & vbcrlf
       EventType = EVENT_ERROR
       Err.Clear
    Else
       RunningLog = RunningLog & "Drive mapped succesfully." & vbcrlf
    End If
    Err.Clear
  Else
    RunningLog = RunningLog & "Drive mapped successfully.  No mapping was present."
  End If
    objShell.LogEvent EventType, RunningLog
  On Error GoTo 0
End Function


Function IsMember(strGroup)
' Function to test for user group membership.
' strGroup is the NT name (sAMAccountName) of the group to test.
' objGroupList is a dictionary object, with global scope.
' Returns True if the user is a member of the group.

  If IsEmpty(objGroupList) Then
    Call LoadGroups
  End If
  IsMember = objGroupList.Exists(strGroup)
End Function


Sub LoadGroups
' Subroutine to populate dictionary object with group memberships.
' objUser is the user object, with global scope.
' objGroupList is a dictionary object, with global scope.

  Dim objGroup
  Set objGroupList = CreateObject("Scripting.Dictionary")
  objGroupList.CompareMode = vbTextCompare
  For Each objGroup In objUser.Groups
    objGroupList(objGroup.name) = True
  Next
  Set objGroup = Nothing
End Sub

Sub RemoveAllNetworkDrives
  On Error Resume Next
  objShell.LogEvent EVENT_INFO, strScriptName & " is removing all network drives." 
  Dim objDrive,intDrive
  Set objDrive = objNetwork.EnumNetworkDrives()
  For intDrive = 0 to objDrive.Count -1 Step 2
  objNetwork.RemoveNetworkDrive objDrive.Item(intDrive), True, True
  Next
  On Error Goto 0
End Sub
