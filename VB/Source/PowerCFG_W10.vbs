Option Explicit


Call PowerCFG(Main)



Function PowerCFG(ByRef Main)

	On Error Resume Next

	Dim objWshShell
	Dim regTimeout
	Dim strTimeout
	Dim strPowerScheme
	Dim strCommand
	Dim strComputer
	Dim objWMIService
	Dim colItems 
	Dim objItem
	Dim strmodel
	strComputer = "."
	Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2")
	Set colItems = objWMIService.ExecQuery("SELECT model FROM Win32_ComputerSystem")
	For Each objItem In colItems
	strmodel =  objItem.Model
	Next

	Set objWshShell = WScript.CreateObject("WScript.Shell")
	strPowerScheme = "NomuraPowerPlan"
	regTimeout = "HKCU\Software\Policies\Microsoft\Windows\Control Panel\Desktop\ScreenSaveTimeOut"
	strTimeout = objWshShell.RegRead(regTimeout)
	
	If strTimeout = "" Then
		strTimeout = 0
	Else
		strTimeout = (strTimeout + 300) / 60
	End If
	
	
	'Begin building log content
	Main.LogFile.Log "PowerCFG: Initialized", True
	Main.LogFile.Log "PowerCFG: JobType " & Ucase(Main.JobType), True
	Main.LogFile.Log "PowerCFG: About to set Power Scheme", True

	if (instr(strmodel,"VMware Virtual Platform") or instr(strmodel,"VMware7,1") or instr(strmodel,"Virtual Machine"))then
		Main.Logfile.Log "PowerCFG: No monitor time-out for VM ", True
		strCommand = "PowerCFG.exe -change -monitor-timeout-ac 0"
		Call LaunchCommand(strCommand)
	Else
		If Ucase(Main.JobType) = "LOGON" Then
			Main.Logfile.Log "PowerCFG: Effective monitor time-out value is " & strTimeout, True
			strCommand = "PowerCFG.exe -change -monitor-timeout-ac " & strTimeout
			Call LaunchCommand(strCommand)
		Else
			Main.Logfile.Log "PowerCFG: Effective monitor time-out value revert to 20mins", True
			strCommand = "PowerCFG.exe -change -monitor-timeout-ac 20"
			Call LaunchCommand(strCommand)
		End If
	End If
	
	Main.LogFile.Log "PowerCFG: Completed", False
	
End Function

Function LaunchCommand(strCommand)
	Dim objWshShell
	Dim intReturn
	
	Set objWshShell = WScript.CreateObject("WScript.Shell")
	
	Main.LogFile.Log "PowerCFG: About to launch command '" & strCommand & "'", True
	'Launch the required script
	intReturn = objWshShell.Run(strCommand, 0, True)
	
	If (Err.Number = 0) Then
		If (intReturn = 0) Then
			Main.LogFile.Log "PowerCFG: Command run successfully", True
		Else 
			Main.LogFile.Log "PowerCFG: Command completed with error(s). Exit code '" & intReturn & "'", True
		End If 
	Else
		Main.LogFile.Log "PowerCFG: Error running command. " & Err.Description & " (" & Err.Number & ")", True
	End If
End Function
