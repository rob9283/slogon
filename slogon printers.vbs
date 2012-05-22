'**********************************************************************************
'**********************************************************************************
'**********************************************************************************
' 
' slogon printers.vbs
'
' A vbscript to map shared folders at logon using simple commands.
' Each drive mapping invoked below records an entry in the local Application log.
' 
' Created by Rob Pennoyer May , 2006.
' 
' Instructions below.  
'
'*** It's helpful to enable "Point and print restrictions" in group policy in
'*** both Computer\Policies\Admin\Printers and User\Policies\Admin\Control\Printers
'*** and to set both "Do not show warning or elevation prompt" items, and make 
'*** sure the policy applies to both users and computers
'
'**********************************************************************************
'**********************************************************************************
'**********************************************************************************


Option Explicit
Dim objNetwork, objShell, objComputer, objADSysInfo
Dim strNetBIOSDomain, strComputername, strCurrentDefaultPrinter
Dim strScriptName, strDCSiteName
Set objNetwork = CreateObject("WScript.Network")
Set objShell = CreateObject("WScript.Shell")
Const EVENT_ERROR = 1, EVENT_WARNING = 2, EVENT_INFO = 4
Const REMOVE_NONE = 1, REMOVE_NETWORK = 2, REMOVE_ALL = 4

'****** Record script start in Application event log.
strScriptName = WScript.ScriptName
objShell.LogEvent EVENT_INFO, strScriptName & " script starting." 

'**********************
'**********************
'NetBIOS Domain name -- You must specify this here:

strNetBIOSDomain = "DOMAIN"

'**********************
'**********************

strComputerName = objNetwork.ComputerName
Set objComputer = GetObject("WinNT://" & strNetBIOSDomain & "/" & strComputerName & ",computer")
Set objADSysInfo = CreateObject("ADSystemInfo")
strDCSiteName = objADSysInfo.GetDCSiteName(strComputerName)


'**********************************************************************************
'**********************************************************************************
'**********************************************************************************
' Edit commands within these bars only.  
' ***Don't forget to set the NETBIOS domain name up above.
' There are three commands available for use.  Use them in any combination or
' quantity you choose.  You MUST start with WalkExistingPrinterList.  The rest of 
' the commands are self-explanatory.
' Functions that check group membership check security groups, not OUs.  
' It is recommended to use computer group membership to map printers, not user groups.



WalkExistingPrinterList(REMOVE_NETWORK)
'Arguments:   REMOVE_NETWORK        removes network printers only
'             REMOVE_ALL            removes all printers
'             REMOVE_NONE           leaves all printers



'MapPrinter "\\172.28.110.19\hplj4200"
'MapPrinter "\\server03\Canon ImageRunner 5135C"

'SetDefaultPrinter "\\server01\HP2200"

'RestoreExistingDefaultPrinter

'ComputerSiteMapPrinter "NYC", "\\server01\hp5660"
'ComputerSiteSetDefaultPrinter "SecondFloor", "\\server01\colorlaser"
'ComputerSiteRestoreExistingDefaultPrinter


'ComputerGroupMapPrinter "executivePCs","\\server04\HP5660"
'ComputerGroupSetDefaultPrinter "accounting", "\\server04\hp5660"
'ComputerGroupRestoreExistingDefaultPrinter


'UserGroupMapPrinter "autocadusers","\\server05\plotter"
'UserGroupSetDefaultPrinter "autocadusers","\\server05\plotter"
'UserGroupRestoreExistingDefaultPrinter


'**********************************************************************************
'**********************************************************************************
'**********************************************************************************

'****** Clean up.
objShell.LogEvent EVENT_INFO, strSCriptName & " script complete." 
Set objNetwork = Nothing
Set objShell = Nothing
Set objComputer = Nothing

Function WalkExistingPrinterList(RemoveWhat)

     Dim strComputer, objWMIService, colInstalledPrinters, objPrinter
     Dim strRunningLog, EventType
     EventType = EVENT_INFO
     strRunningLog = strScriptName & " is walking the printer list and will "
     if RemoveWhat = REMOVE_NONE Then
       strRunningLog = strRunningLog & "not remove any printers." & vbcrlf
       Else If RemoveWhat = REMOVE_NETWORK Then
         strRunningLog = strRunningLog & "attempt to remove network printers only." & vbcrlf
           Else If RemoveWhat = REMOVE_ALL Then
             strRunningLog = strRunningLog & "attempt to remove all printers." & vbcrlf
           End If
       End If
     End If
     strComputer = "."
     Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
     Set colInstalledPrinters = objWMIService.ExecQuery ("Select * from Win32_Printer")
     For Each objPrinter in colInstalledPrinters
     if objPrinter.default = True then 
          strCurrentDefaultPrinter = objPrinter.Name
          strRunningLog = strRunningLog & "Current Default Printer is " & strCurrentDefaultPrinter & vbcrlf
     end if

     if RemoveWhat = REMOVE_NETWORK Then
         if InStr(objPrinter.Name, "\\") > 0 then
               strRunningLog = strRunningLog & "Attempting to remove printer: " & objPrinter.Name & vbcrlf
               Err.Clear
               On Error Resume Next
	       objNetwork.RemovePrinterConnection objPrinter.Name
               If Err.Number <> 0 Then
                    strRunningLog = strRunningLog & "Unable to remove printer!" & vbcrlf
                    EventType = EVENT_ERROR
                    Err.Clear
               Else
                    strRunningLog = strRunningLog & "Printer removed successfully!" & vbcrlf
               End If
               Err.Clear
               On Error Goto 0
         End If
     End If
     
     If RemoveWhat = REMOVE_ALL Then
          strRunningLog = strRunningLog & "Attempting to remove printer: " & objPrinter.Name & vbcrlf
          Err.Clear
          On Error Resume Next
          objNetwork.RemovePrinterConnection objPrinter.Name
          If Err.Number <> 0 Then
               strRunningLog = strRunningLog & "Unable to remove printer!" & vbcrlf
               EventType = EVENT_ERROR
               Err.Clear
          Else
               strRunningLog = strRunningLog & "Printer removed successfully!" & vbcrlf
          End If
          Err.Clear
          On Error Goto 0
     End If

     Next
     objShell.LogEvent EventType, strRunningLog

End Function


'****** Computer Group Functions
Function ComputerGroupRestoreExistingDefaultPrinter(strGroup)
     If IsComputerMember(strGroup) Then
          RestoreExistingDefaultPrinter
     End If
End Function

Function ComputerGroupMapPrinter(strGroup, strPrinter)
     If IsComputerMember(strGroup) Then
          MapPrinter(strPrinter)
     End If
End Function

Function ComputerGroupSetDefaultPrinter(strGroup, strPrinter)
     If IsComputerMember(strGroup) Then
          SetDefaultPrinter(strPrinter)
     End If
End Function
'******


'****** Computer Site Functions
Function ComputerSiteRestoreExistingDefaultPrinter(strSite)
     If IsComputerInSite(strSite) Then
          RestoreExistingDefaultPrinter
     End If
End Function

Function ComputerSiteMapPrinter(strSite, strPrinter)
     If IsComputerInSite(strSite) Then
          MapPrinter(strPrinter)
     End If
End Function

Function ComputerSiteSetDefaultPrinter(strSite, strPrinter)
     If IsComputerInSite(strSite) Then
          SetDefaultPrinter(strPrinter)
     End If
End Function
'******



'****** User Group Functions
Function UserGroupRestoreExistingDefaultPrinter(strGroup)
     If IsUserMember(strGroup) Then
           RestoreExistingDefaultPrinter
     End If
End Function

Function UserGroupMapPrinter(strGroup, strPrinter)
     If IsUserMember(strGroup) Then
          MapPrinter(strPrinter)
     End If
End Function

Function UserGroupSetDefaultPrinter(strGroup, StrPrinter)
     If IsComputerMember(strGroup) Then
          SetDefaultPrinter(strPrinter)
     End If
End Function
'******


Sub RestoreExistingDefaultPrinter
     SetDefaultPrinter(strCurrentDefaultPrinter)
End Sub



Function MapPrinter(strPrinter)

     Dim strRunningLog, EventType

     'First sleep for .2 seconds, so the commands don't fall on top of each other
     WScript.sleep 200
     strRunningLog = strScriptName & " is attempting to map the printer:" & vbcrlf & strPrinter & vbcrlf
     EventType = EVENT_INFO
     On Error Resume Next
     Err.Clear
     objNetwork.AddWindowsPrinterConnection strPrinter
     if Err.Number <> 0 Then
        strRunningLog = strRunningLog & "Error mapping printer!  Giving up." & vbcrlf
        EventType = EVENT_ERROR
        Err.Clear
     Else
        strRunningLog = strRunningLog & "Printer mapped successfully!" & vbcrlf
     End If
     Err.Clear
     objShell.LogEvent EventType, strRunningLog
     On Error Goto 0

End Function



Function SetDefaultPrinter(strPrinter)

     Dim strRunningLog, EventType

     'First sleep for .2 seconds, so the commands don't fall on top of each other
     WScript.sleep 200
     strRunningLog = strScriptName & " is attempting to set as default printer:" & vbcrlf & strPrinter & vbcrlf
     EventType = EVENT_INFO
     On Error Resume Next
     Err.Clear
     objNetwork.SetDefaultPrinter strPrinter
     if Err.Number <> 0 Then
        strRunningLog = strRunningLog & "Error mapping printer!  Giving up." & vbcrlf
        EventType = EVENT_ERROR
        Err.Clear
     Else
        strRunningLog = strRunningLog & "Printer successfully set as default!" & vbcrlf
     End If
     Err.Clear
     objShell.LogEvent EventType, strRunningLog
     On Error Goto 0

End Function



Function IsComputerMember(strGroup)
    Dim objGroup
    Set objGroup = GetObject("WinNT://" & strNetBIOSDomain & "/" & strGroup & ",group")
    If objGroup.IsMember(objComputer.ADsPath & "$") Then
        IsComputerMember = True
    Else
        IsComputerMember = False
    End If
    Set objGroup = Nothing
End Function



Function IsUserMember(strGroup)
' Function to test for user group membership.
' strGroup is the NT name (sAMAccountName) of the group to test.
' objGroupList is a dictionary object, with global scope.
' Returns True if the user is a member of the group.

  If IsEmpty(objGroupList) Then
    Call LoadGroups
  End If
  IsUserMember = objGroupList.Exists(strGroup)
End Function


Function IsComputerInSite(strSite)
' Function to test if computer is in a given site
   if strSite = strDCSiteName Then
		IsComputerInSite = True
	Else
		IsComputerInSite = False
	End If
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





