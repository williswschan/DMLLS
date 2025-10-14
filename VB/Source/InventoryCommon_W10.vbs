'This module contains classes and functions that are shared by Inventory modules

Option Explicit


Class InventoryServer
	Dim ServerName
	Dim Online
	Dim ResponseTime
	Dim ServiceUrl
	Dim ServiceAvailable
	
	Sub Class_Initialize()
		ServerName = ""
		Online = False
		ResponseTime = -1
		ServiceUrl = ""
		ServiceAvailable = False
	End Sub
	
	Function ToString
		ToString = _
			"Inventory SERVER:" & vbCrLf & _
			"Server Name:       " & ServerName & vbCrLf &_
			"Online:            " & CStr(Online) & vbCrLf &_
			"Ping Response:     " & ResponseTime & " ms" & vbCrLf &_
			"Service URL:       " & ServiceUrl & vbCrLf &_
			"Service Available: " & CStr(ServiceAvailable)
	End Function
End Class

'Find best available InventoryServer for the computer
Function GetInventoryServer(strComputerDomain)
	Dim objAdSysInfo, strServerName, strServiceName, strServerFqdn
	Dim objInventoryServer
	Dim colPing, objPingStatus
	Dim objHttp, strTestRequest
	Dim intTimeOut
	
	On Error Resume Next
	
	strServerName 	= Trim(getResource("InventoryServer"))
	strServiceName 	= Trim(getResource("InventoryService"))
	intTimeOut 	= Trim(getResource("InventoryServiceTimeout"))

	Set objAdSysInfo = CreateObject("ADSystemInfo")
	If (UCase(objAdSysInfo.ForestDNSName) = "QA.NOM") Then
		strServerFqdn = "gdpmappercbqa.nomura.com"
	Else
		strServerFqdn = "gdpmappercb.nomura.com"
	End If
	
	strTestRequest = _
		"<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" &_
		" <soap:Body>" &_
		"  <TestService xmlns=""http://webtools.japan.nom"" />" &_
		" </soap:Body>" &_
		"</soap:Envelope>"
	
	Set objInventoryServer = New InventoryServer

	objInventoryServer.ServerName = strServerFqdn
	objInventoryServer.ServiceURL = "http://" & strServerFqdn & "/" & strServiceName
		
	Set colPing = GetObject("winmgmts:").ExecQuery("select * from Win32_PingStatus where address = '" & strServerFqdn & "'")

	If (Err.Number = 0) Then
		For Each objPingStatus in colPing
			If (objPingStatus.StatusCode = 0) Then
				objInventoryServer.Online = True
				objInventoryServer.ResponseTime = objPingStatus.ResponseTime
			Else
				objInventoryServer.Online = False
			End If
		Next
	Else
		Err.Clear
	End If

	If (objInventoryServer.Online) Then
		Set objHttp = CreateObject("Msxml2.XMLHTTP.3.0")
			
		objHttp.Open "POST", objInventoryServer.ServiceURL, True
		objHttp.SetRequestHeader "Content-Type", "text/xml; CharSet=UTF-8"
			
		objHttp.Send strTestRequest
			
		'Async request was made. Need to wait until data has all been returned by the server (readystate = 4). The timeout is not 
		Do Until ((objHttp.ReadyState = 4) Or (intTimeOut <= 0))
			intTimeOut = intTimeOut - 100
			Wscript.Sleep 100
		Loop
				
		If (Err.Number = 0) Then
			If (intTimeOut > 0) Then
				If (objHttp.Status = 200) Then
					'Expect response text to only contain "True" or "False"
					If (StrComp(Trim(objHttp.ResponseXML.Text), "True", 1) = 0) Then
						objInventoryServer.ServiceAvailable = True						
					End If
				End If
			End If
		Else
			Err.Clear
		End If
		
		If (objInventoryServer.ServiceAvailable) Then
			'Return most responsive and available InventoryServer and exit Function
			Set GetInventoryServer = objInventoryServer
			Exit Function
		End If
	Else
		'No available service was found
		Set GetInventoryServer = Nothing
	End If	
	
	On Error Goto 0
	
End Function

Function EscapeXMLText(strText)
	'Guard against null value for paramater
	If (IsNull(strText)) Then
		EscapeXMLText = ""
		Exit Function
	End If
	
	'Replace special characters
	strText = Replace(strText, "&", "&amp;")
	strText = Replace(strText, "'", "&apos;")
	strText = Replace(strText, """", "&quot;")
	strText = Replace(strText, "<", "&lt;")
	strText = Replace(strText, ">", "&gt;")
	
	EscapeXMLText = strText
End Function

Function IsUserPartOfGroup(strGroupName, intLeftIndex)

	Dim retVal
	Dim adGroup

	retVal = False

	If (Not IsNull(Main.User.Groups)) Then

		For Each adGroup In Main.User.Groups

		   If (StrComp(strGroupName , Left(adGroup,intLeftIndex) , 1) = 0) Then
			    retVal = True
		   Else
		   		retVal = False
		   End If
		Next

	End If

	IsUserPartOfGroup = retVal

End Function

Function IsOfflineLaptopInventory(OfflineLaptopGroup)

	Dim retVal
	Dim adGroup

	retVal = False

	If (Not IsNull(Main.User.Groups)) Then

		For Each adGroup In Main.User.Groups

		   If InStr(adGroup,OfflineLaptopGroup) > 0 Then
			    retVal = True
		   End If
		Next

	End If

	IsOfflineLaptopInventory = retVal

End Function

Function IsCiscoVPNConnected()
	Dim objWMIService, ObjItem
	Dim  strComputer, colItems
	Dim retVal
	'On Error Resume Next
	strComputer = "."

	Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
	Set colItems = objWMIService.ExecQuery 	("Select * from Win32_NetworkAdapter",,48)

	retVal=false

	For Each objItem in colItems

	If objItem.MACAddress <> "" and Instr(Ucase(objItem.Name),"CISCO") > 0 Then
  		if objItem.Netconnectionstatus = "2" then
     			retVal=true 
  		else
     			retVal=false
  		end if 
   
	End IF

	Next
	IsCiscoVPNConnected=retVal
End Function