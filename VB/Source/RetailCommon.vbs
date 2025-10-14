'==========================================================================
'
' NAME: RetailCommon.vbs
' 
' VERSIONS:
'-------------------------------------------------------------
' 	VERSION	: 1.0
' 	AUTHOR	: Nitish Kolwankar / Toru Ozawa
'	DATE  	: 05/08/2015 / 02/12/2023
' 	COMMENT	: 1st Version for Retail GDP Migration
'                 Nomura Trust Band added
'-------------------------------------------------------------
'
'==========================================================================

Function IsRetailUser

	Dim retVal

	retVal = False
	
	If (InStr(1,Main.User.DN,"OU=Nomura Retail") > 0) or (InStr(1,Main.User.DN,"OU=Nomura Trust Bank") > 0) or (InStr(1,Main.User.DN,"OU=TOK,OU=Nomura Asset Management") > 0) Then
		retVal = True
	End If	
	
	IsRetailUser = retVal
	
End Function

Function IsRetailHost
	
	Dim retVal

	retVal = False
	
	If (InStr(1,Main.Computer.DN,"OU=Nomura Retail") > 0) or (InStr(1,Main.Computer.DN,"OU=Nomura Trust Bank") > 0) or (InStr(1,Main.Computer.DN,"OU=TOK,OU=Nomura Asset Management") > 0)Then
		retVal = True
	End If	
	
	IsRetailHost = retVal
	
End Function

Function IsRetailUserPartOfGroup(strGroupName)

	Dim retVal
	Dim objWshShell
	Dim objGroup

	retVal = False

	Set objWshShell = CreateObject("WScript.Shell")

	Set objGroup = objWshShell.exec("whoami /groups")

	If instr(Ucase(objGroup.stdOut.ReadAll),Ucase(strGroupName)) = 0 then
	 retVal = False
	else
	 retVal = True
	end if

	Set objGroup = nothing

	IsRetailUserPartOfGroup = retVal

End Function


Function IsSharedVDI()

	Dim strHostName
	IsSharedVDI = False	

	strHostName = Main.Computer.Name
	
	If (InStr(1, strHostName, "JPRWV1") > 0 OR InStr(1, strHostName, "JPRWV3") > 0) Then
		IsSharedVDI =  True
	End If

End Function
