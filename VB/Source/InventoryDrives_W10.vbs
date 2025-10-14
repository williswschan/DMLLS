Option Explicit


If IsCiscoVPNConnected() Then
	Main.LogFile.Log "Inventory Drive: Identified as VPN connected", True
	Main.LogFile.Log "Inventory Drive: VPN connected Laptop is requested to skip this process", True
	Main.LogFile.Log "Inventory Drive: Completed.", True

Else
	Main.LogFile.Log "Inventory Drive: Identified as VPN not connected", True
	Call GatherAndInsertDriveInventory(Main)

End If

Class InventoryDrive
	Dim Id
	Dim UserId
	Dim HostName
	Dim Domain
	Dim SiteName
	Dim City
	Dim Drive
	Dim UncPath
	Dim Description
	Dim LastUpdate
	Dim OuMapping

	
	Sub Class_Initialize()
		Id = vbNull
		UserId = vbNull
		HostName = vbNull
		Domain = vbNull
		SiteName = vbNull
		City = vbNull
		Drive = vbNull
		UncPath = vbNull
		Description = ""
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

				Case "HOSTNAME"
					HostName = xmlNode.Text

				Case "DOMAIN"
					Domain = xmlNode.Text
					
				Case "SITENAME"
					SiteName = xmlNode.Text
					
				Case "CITY"
					City = xmlNode.Text

				Case "DRIVE"
					Drive = xmlNode.Text
					
				Case "UNCPATH"
					UncPath = xmlNode.Text
					
				Case "DESCRIPTION"
					Description = xmlNode.Text

				Case "LASTUPDATE"
					LastUpdate = xmlNode.Text

				Case "OUMAPPING"
					OuMapping = xmlNode.Text
			End Select
		Next
	End Function
	
	
	Function ToString
		ToString = _
			"INVENTORY DRIVE:" & vbCrLf & _
			"Id: " & Id & vbCrLf &_
			"UserId: " & UserId & vbCrLf &_
			"HostName: " & HostName & vbCrLf &_
			"Domain: " & Domain & vbCrLf &_
			"SiteName: " & SiteName & vbCrLf &_
			"City: " & City & vbCrLf &_
			"Drive: " & Drive & vbCrLf &_
			"UncPath: " & UncPath & vbCrLf &_
			"Description: " & Description & vbCrLf &_
			"LastUpdate: " & LastUpdate &_
			"OuMapping: " & OuMapping
	End Function


	Function ToXml
		ToXml = _
			"<InventoryDrive>" & _
			"<UserId>" & UserId & "</UserId>" & _
			"<HostName>" & HostName & "</HostName>" & _
			"<Domain>" & Domain & "</Domain>" & _
			"<SiteName>" & SiteName & "</SiteName>" & _
			"<City>" & City & "</City>" & _
			"<Drive>" & Drive & "</Drive>" & _
			"<UncPath>" & UncPath & "</UncPath>" & _
			"<Description>" & Description & "</Description>" & _
			"<OuMapping>" & OuMapping & "</OuMapping>" & _
			"</InventoryDrive>"
	End Function
End Class


Class InventoryDrives
	Dim Drives()
	Dim DriveCount

	
	Sub Class_Initialize()
		DriveCount = 0
	End Sub


	Function ToString
		Dim retVal
		Dim drive
		
		retVal = "INVENTORY DRIVES:"
		
		For Each drive In Drives
			retVal = retVal & vbCrLf & Replace(drive.ToString, vbCrLf, vbCrLf & vbTab)
		Next
		
		ToString = retVal
	End Function


	Function ToXml
		Dim retVal
		Dim drive
		
		retVal = "<Mappings>"
		
		For Each drive In Drives
			retVal = retVal & vbCrLf & Replace(drive.ToXml, vbCrLf, vbCrLf & vbTab)
		Next
		
		ToXml = retVal & "</Mappings>"
	End Function


	Function AddDrive(UserId, HostName, Domain, OuMapping, SiteName, City, Drive, UncPath, Description)
		Dim invDrive

		Set invDrive = New InventoryDrive

		invDrive.UserId = UserId
		invDrive.HostName = HostName
		invDrive.Domain = Domain
		invDrive.SiteName = SiteName
		invDrive.City = City
		invDrive.Drive = Drive
		invDrive.UncPath = UncPath
		invDrive.Description = Description
		invDrive.OuMapping = OuMapping

		If (DriveCount = 0) Then
			ReDim Drives(0)
		Else
			ReDim Preserve Drives(UBound(Drives) + 1)
		End If

		Set Drives(DriveCount) = invDrive

		DriveCount = DriveCount + 1
	End Function
End Class


Function InsertDriveInventory(ByRef Main, Drives)
	Dim inventoryServer
	Dim strRequest
	Dim http
	Dim intTimeOut

	Main.LogFile.Log "Inventory Drive: About to insert drive mapping data to the service", False
	
	intTimeOut = Trim(getResource("InventoryServiceTimeout"))
	
	Main.LogFile.Log "Inventory Drive: Timeout value set as '" & intTimeOut & "' ms", False

	Set inventoryServer = Nothing	
	strRequest = ""

	Set inventoryServer = GetInventoryServer(Main.Computer.Domain)		
	
	If (inventoryServer Is Nothing) Then
		Main.LogFile.Log "Inventory Drive: No available inventory service was found", True
		Main.LogFile.Log "Inventory Drive: Completed", True

		Exit Function
	End If

	Main.LogFile.Log "Inventory Drive: Best available inventory service is '" & inventoryServer.ServiceURL & "'", True
	Main.LogFile.Log "Inventory Drive: Server Selection" & vbCrLf & vbCrLf & inventoryServer.ToString & vbCrLf, False

	strRequest = _
		"<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" &_
		" <soap:Body>" &_
		"  <InsertActiveDriveMappingsFromInventory xmlns=""http://webtools.japan.nom"">" &_
		Drives.ToXml &_
		"  </InsertActiveDriveMappingsFromInventory>" &_
		" </soap:Body>" &_
		"</soap:Envelope>"

	Main.LogFile.Log "Inventory Drive: Request content is " & vbCrLf & vbCrLf & strRequest & vbCrLf, False

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
			Main.LogFile.Log "Inventory Drive: Exceeded timeout for web service", True
		End If
	Loop
	
	If (Err.Number = 0) Then
		If (intTimeOut > 0) Then
			If (http.Status = 200) Then
				Main.LogFile.Log "Inventory Drive: Received response from web service " & vbCrLf & vbCrLf & http.ResponseText & vbCrLf, False
			Else
				Main.LogFile.Log "Inventory Drive: HTTP request status was not 200. Drive inventory could not be inserted (Status: " & http.Status & ", Response: " & http.ResponseText & ")", True
			End If
		End If
	Else
		Main.LogFile.Log "Inventory Drive: inventory service unavailable or not responding", True

		Exit Function
	End If

	On Error Goto 0
End Function


Function GatherAndInsertDriveInventory(ByRef Main)
	Dim networkDrives
	Dim networkDrive
	Dim inventoryDrives
	Dim description
	Dim itt
	Dim computerName
	Dim key
	Dim rootKey
	Dim driveDescription
	Dim valueName
	Dim regObject
	Dim subKey
	Dim strHomeShare

	Main.LogFile.Log "Inventory Drive: Gathering and inserting drives", False

	Set inventoryDrives = New InventoryDrives

	networkDrives = GetMappedDrives

	If IsRetailUser <> 0 Then
		computerName = "."
		rootKey = HKEY_CURRENT_USER
		key = "Volatile Environment"
		valueName = "HOMEDRIVE"
		driveDescription = ""
		
		On Error Resume Next 
		Set regObject=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" &_ 
		   computerName & "\root\default:StdRegProv")
		regObject.GetStringValue rootKey, key,valueName,driveDescription

		If Err.Number <> 0 Then
			Main.LogFile.Log "Inventory Drive: V Drive does not exist for the user.", True
			Exit Function
		Else
			regObject.GetStringValue rootKey, key & "\" & subKey,"HOMESHARE", strHomeShare
			inventoryDrives.AddDrive _
				Main.User.Name,_
				Main.Computer.Name,_
				Main.Computer.Domain,_
				Main.User.OUMapping,_
				Main.Computer.Site,_
				Main.Computer.CityCode,_
				"V",_
				strHomeShare,_
				"Home Drive"
		End If
	Else
		If UBound(networkDrives) = 0 Then
			Main.LogFile.Log "Inventory Drive: No Drives are mapped.", True
			Exit Function
		End If
	End If

	For itt = 0 To UBound(networkDrives) Step 3
		computerName = "."
		rootKey = HKEY_CURRENT_USER
		key = "Software\Microsoft\Windows\CurrentVersion\Explorer\MountPoints2"
		valueName = "_LabelFromReg"
		subKey = Replace(networkDrives(itt + 1),"\","#")
		driveDescription = ""
		
		On Error Resume Next 
		Set regObject=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" &_ 
		   computerName & "\root\default:StdRegProv")
		regObject.GetStringValue rootKey, key & "\" & subKey,valueName,driveDescription

		If Err.Number <> 0 Then
			driveDescription = ""
		End If 
		
		inventoryDrives.AddDrive _
			Main.User.Name,_
			Main.Computer.Name,_
			Main.Computer.Domain,_
			Main.User.OUMapping,_
			Main.Computer.Site,_
			Main.Computer.CityCode,_
			networkDrives(itt),_
			networkDrives(itt + 1),_
			driveDescription
	Next

	Main.LogFile.Log "Inventory Drive: Inserting " & inventoryDrives.DriveCount & " drives", False

	InsertDriveInventory Main, inventoryDrives
End Function