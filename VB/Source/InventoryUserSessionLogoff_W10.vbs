Option Explicit

Call InsertLogonSession(Main)

Class InventoryUserSession
	Dim Id
	Dim UserId
	Dim UserDomain
	Dim HostName
	Dim Domain
	Dim SiteName
	Dim City
	Dim LastUpdate
	Dim OuMapping

	
	Sub Class_Initialize()
		Id = vbNull
		UserId = vbNull
		HostName = vbNull
		Domain = vbNull
		SiteName = vbNull
		City = vbNull
		LastUpdate = vbNull
		OuMapping = vbNull
	End Sub
	
	
	Function SetupFromXml(Xml)
		Dim xmlNode
		Dim xmlNodes
		
		Set xmlNodes = Xml.SelectNodes("//InventoryDrive/*")
		
		For Each xmlNode In xmlNodes
			Select Case UCase(xmlNode.baseName)
				Case "ID"
					Id = xmlNode.Text
					
				Case "USERID"
					UserId = xmlNode.Text

				Case "USERDOMAIN"
					UserId = xmlNode.Text

				Case "HOSTNAME"
					HostName = xmlNode.Text

				Case "DOMAIN"
					Domain = xmlNode.Text
					
				Case "SITENAME"
					SiteName = xmlNode.Text
					
				Case "CITY"
					City = xmlNode.Text

				Case "LASTUPDATE"
					LastUpdate = xmlNode.Text
					
				Case "OUMAPPING"
					OuMapping = xmlNode.Text
			End Select
		Next
	End Function
	
	
	Function ToString
		ToString = _
			"INVENTORY USER SESSION:" & vbCrLf & _
			"Id: " & Id & vbCrLf &_
			"UserId: " & UserId & VbCrLf &_
			"UserDomain: " & UserDomain & vbCrLf &_
			"HostName: " & HostName & vbCrLf &_
			"Domain: " & Domain & vbCrLf &_
			"SiteName: " & SiteName & vbCrLf &_
			"City: " & City & vbCrLf &_
			"LastUpdate: " & LastUpdate &_
			"OuMapping: " & OuMapping
	End Function


	Function ToXml
		ToXml = _
			"<InventoryUserSession>" & _
			"<UserId>" & UserId & "</UserId>" & _
			"<UserDomain>" & UserDomain & "</UserDomain>" & _
			"<HostName>" & HostName & "</HostName>" & _
			"<Domain>" & Domain & "</Domain>" & _
			"<SiteName>" & SiteName & "</SiteName>" & _
			"<City>" & City & "</City>" & _
			"<OuMapping>" & OuMapping & "</OuMapping>" & _
			"</InventoryUserSession>"
	End Function
End Class


Class InventoryUserSessions
	Dim UserSessions()
	Dim UserSessionCount

	
	Sub Class_Initialize()
		UserSessionCount = 0
	End Sub


	Function ToString
		Dim retVal
		Dim userSession
		
		retVal = "INVENTORY USER SESSIONS:"
		
		For Each userSession In UserSessions
			retVal = retVal & vbCrLf & Replace(userSession.ToString, vbCrLf, vbCrLf & vbTab)
		Next
		
		ToString = retVal
	End Function


	Function ToXml
		Dim retVal
		Dim userSession
		
		retVal = "<Sessions>"
		
		For Each userSession In UserSessions
			retVal = retVal & vbCrLf & Replace(userSession.ToXml, vbCrLf, vbCrLf & vbTab)
		Next
		
		ToXml = retVal & "</Sessions>"
	End Function


	Function AddUserSession(UserId, UserDomain, HostName, Domain, SiteName, City, OuMapping)
		Dim invUserSession

		Set invUserSession = New InventoryUserSession

		invUserSession.UserId = UserId
		invUserSession.UserDomain = UserDomain
		invUserSession.HostName = HostName
		invUserSession.Domain = Domain
		invUserSession.SiteName = SiteName
		invUserSession.City = City
		invUserSession.OuMapping = OuMapping

		If (UserSessionCount = 0) Then
			ReDim UserSessions(0)
		Else
			ReDim Preserve UserSessions(UBound(UserSessions) + 1)
		End If

		Set UserSessions(UserSessionCount) = invUserSession

		UserSessionCount = UserSessionCount + 1
	End Function
End Class


Function InsertLogonInventory(ByRef Main, UserSessions)
	Dim inventoryServer
	Dim strRequest
	Dim http
	Dim intTimeOut

	Main.LogFile.Log "Inventory Logon: About to insert user session data to the service", False
	
	intTimeOut = Trim(getResource("InventoryServiceTimeout"))
	
	Main.LogFile.Log "Inventory Logon: Timeout value set as '" & intTimeOut & "' ms", False

	strRequest = ""

	Set inventoryServer = GetInventoryServer(Main.Computer.Domain)

	If (inventoryServer Is Nothing) Then
		Main.LogFile.Log "Inventory Logon: No available inventory service was found", True
		Main.LogFile.Log "Inventory Logon: Completed", True

		Exit Function
	End If

	Main.LogFile.Log "Inventory Logon: Best available inventory service is '" & inventoryServer.ServiceURL & "'", True
	Main.LogFile.Log "Inventory Logon: Server Selection" & vbCrLf & vbCrLf & inventoryServer.ToString & vbCrLf, False

	strRequest = _
		"<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" &_
		" <soap:Body>" &_
		"  <InsertLogoffInventory xmlns=""http://webtools.japan.nom"">" &_
		UserSessions.ToXml &_
		"  </InsertLogoffInventory>" &_
		" </soap:Body>" &_
		"</soap:Envelope>"

	On Error Resume Next
	
	Set http = CreateObject("Msxml2.XMLHTTP.3.0")
	
	http.Open "POST", inventoryServer.ServiceURL, false
	http.SetRequestHeader "Content-Type", "text/xml; CharSet=UTF-8"
		
	http.Send strRequest

	'Async request was made. Need to wait until data has all been returned by the server (readystate = 4). The timeout is not 
	Do Until ((http.ReadyState = 4) Or (intTimeOut <= 0))
		intTimeOut = intTimeOut - 100

		Wscript.Sleep 100
		
		If (intTimeOut <= 0) Then
			Main.LogFile.Log "Inventory Logon: Exceeded timeout for web service", True
		End If
	Loop
	
	If (Err.Number = 0) Then
		If (intTimeOut > 0) Then
			If (http.Status = 200) Then
				Main.LogFile.Log "Inventory Logon: Received response from web service " & vbCrLf & vbCrLf & http.ResponseText & vbCrLf, False
			Else
				Main.LogFile.Log "Inventory Logon: HTTP request status was not 200. User session inventory could not be inserted (Status: " & http.Status & ", Response: " & http.ResponseText & ")", True
			End If
		End If
	Else
		Main.LogFile.Log "Inventory Logon: inventory service unavailable or not responding", True

		Exit Function
	End If

	On Error Goto 0
End Function

Function InsertLogonSession(ByRef Main)
	Dim InventoryUserSessions
	
	Set InventoryUserSessions = New InventoryUserSessions
	
	InventoryUserSessions.AddUserSession  _
			Main.User.Name,_
			Main.User.Domain,_
			Main.Computer.Name,_
			Main.Computer.Domain,_
			Main.Computer.Site,_
			Main.Computer.CityCode,_
			Main.User.OuMapping
	
	InsertLogonInventory Main, InventoryUserSessions
	
End Function 