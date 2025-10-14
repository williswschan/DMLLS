' ManageIEZones.vbs v1.5
' 11/02/2011 - Ronald Wong
' 08/08/2011 - Kenjie Mojar
' 29/05/2021 - Calvin Chen ismemberof() performance fix
' 09/02/2022 - Calvin Chen ismemberof() log fix 
Option Explicit

Call ManageIEZones(Main)


Function ManageIEZones(ByRef Main)
	Dim objFSO, objWshShell
	Dim strRegFileName, strRegFileServerPath, strRegFilePath, strCommand
	Dim strDebugGroup, strPilotGroup
	Dim strIEZoneFolder
	Dim intReturn
	
	On Error Resume Next
	
	Main.LogFile.Log "Manage IE Zones: Manage IE zone configuration and assign sites and domains to zones", True
	
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objWshShell = CreateObject("WScript.Shell")
	

	Select Case LCase(Main.JobType)
		Case "logon"
		
			Select Case UCase(Main.User.ShortDomain)
				Case "AMERICAS", "QAAMERICAS", "RNDAMERICAS"
					strDebugGroup = "Desktop Management Script Debug-AM-U"
					strPilotGroup = "Pilot Desktop Management Script-AM-U"
					
				Case "ASIAPAC", "QAASIAPAC", "RNDASIAPAC"
					strDebugGroup = "Desktop Management Script Debug-AP-U"
					strPilotGroup = "Pilot Desktop Management Script-AP-U"
					
				Case "EUROPE", "QAEUROPE", "RNDEUROPE"
					strDebugGroup = "Desktop Management Script Debug-EU-U"
					strPilotGroup = "Pilot Desktop Management Script-EU-U"
				
				Case "JAPAN", "QAJAPAN", "RNDJAPAN"
					strDebugGroup = "Desktop Management Script Debug-JP-U"
					strPilotGroup = "Pilot Desktop Management Script-JP-U"
					
				Case Else
					strLogContent = strLogContent & vbCrLf & "Domain '" & strDomain & "' not recognized"
					strDebugGroup = "Desktop Management Script Debug-JP-U"
					strPilotGroup = "Pilot Desktop Management Script-JP-U"
			End Select
		
			If (IsMemberOf(Main.User.DN, strPilotGroup)) Then
				strIEZoneFolder = "Pilot\"
			Else
				strIEZoneFolder = "Prod\"
			End If
		
			strRegFileName = "IEZones-U.reg"
			strRegFileServerPath = "\\" & Main.User.Domain & "\Apps\ConfigFiles\IEZones\" & strIEZoneFolder & strRegFileName
			
		Case "startup"
		
			Select Case UCase(Main.Computer.ShortDomain)
				Case "AMERICAS", "QAAMERICAS", "RNDAMERICAS"
					strDebugGroup = "Desktop Management Script Debug-AM-C"
					strPilotGroup = "Pilot Desktop Management Script-AM-C"
					
				Case "ASIAPAC", "QAASIAPAC", "RNDASIAPAC"
					strDebugGroup = "Desktop Management Script Debug-AP-C"
					strPilotGroup = "Pilot Desktop Management Script-AP-C"
					
				Case "EUROPE", "QAEUROPE", "RNDEUROPE"
					strDebugGroup = "Desktop Management Script Debug-EU-C"
					strPilotGroup = "Pilot Desktop Management Script-EU-C"
				
				Case "JAPAN", "QAJAPAN", "RNDJAPAN"
					strDebugGroup = "Desktop Management Script Debug-JP-C"
					strPilotGroup = "Pilot Desktop Management Script-JP-C"
					
				Case Else
					strLogContent = strLogContent & vbCrLf & "Domain '" & strDomain & "' not recognized"
					strDebugGroup = "Desktop Management Script Debug-JP-C"
					strPilotGroup = "Pilot Desktop Management Script-JP-C"
			End Select
			
			
		
			If (IsMemberOf(Main.Computer.DN, strPilotGroup)) Then
				strIEZoneFolder = "Pilot\"
			Else
				strIEZoneFolder = "Prod\"
			End If
		
			strRegFileName = "IEZones-M.reg"
			strRegFileServerPath = "\\" & Main.Computer.Domain & "\Apps\ConfigFiles\IEZones\" & strIEZoneFolder & strRegFileName
			
		Case Else
			Main.LogFile.Log  "Manage IE Zones: Script was called via unexpected job type '" & Main.JobType & "'. Will not continue", True
			Main.LogFile.Log  "Manage IE Zones: Completed", True
			Exit Function
	End Select
	
	Main.LogFile.Log "Manage IE Zones: Reg merge file path is '" & strRegFileServerPath & "'", True
	
	'Exit if file cannot be found
	If (objFSO.FileExists(strRegFileServerPath)) Then
		'Create local copy of file to improve performance when parsing (travelling users)
		strRegFilePath = objWshShell.ExpandEnvironmentStrings("%TEMP%\") & strRegFileName
		
		objFSO.CopyFile strRegFileServerPath, strRegFilePath, True
		
		If (Err.Number = 0) Then
			Main.LogFile.Log "Manage IE Zones: Reg merge file copied to '" & strRegFilePath & "'", False
		Else
			'Use the server path
			strRegFilePath = strRegFileServerPath
		End If
	Else
		Main.LogFile.Log "Manage IE Zones: Cannot access the '" & strRegFileServerPath & "' file", True
		Main.LogFile.Log "Manage IE Zones: Completed", True
		Exit Function
	End If
	
	strCommand = "regedit /s """ & strRegFilePath & """"
	Main.LogFile.Log "Manage IE Zones: About to import reg file with command '" & strCommand & "'", False
	
	On Error Resume Next
	
	intReturn = objWshShell.Run(strCommand, 0, True)
	
	If (Err.Number = 0) Then
		If (intReturn = 0) Then
			Main.LogFile.Log "Manage IE Zones: Completed import successfully", True
		Else 
			Main.LogFile.Log "Manage IE Zones: Import completed with error(s). Exit code '" & intReturn & "'", True
		End If 
	Else
		Main.LogFile.Log "Manage IE Zones: Error running reg merge. " & Err.Description & " (" & Err.Number & ")", True
	End If
	
	On Error Goto 0
	
	Main.LogFile.Log "Manage IE Zones: Completed", True
End Function

Function IsMemberOf(strName, strGroupName)
	Dim boolReturn
	Dim objUser, objGroup
	Dim strGroups
	'Default Value
	boolReturn = False
	
	
	
	'Get ADs Object from LDAP ADs provider
	Set  objUser = GetObject("LDAP://" & strName)
	strGroups = Join(objUser.GetEx("memberOf"))

	If InStr(1, strGroups, strGroupName, VbTextCompare) > 0 Then
		boolReturn = True
	' The user is in the group so we can do things
	else
		boolReturn = False
	End If
	
	
	IsMemberOf = boolReturn
End Function

