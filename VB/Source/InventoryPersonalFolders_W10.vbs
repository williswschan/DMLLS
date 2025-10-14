Option Explicit

' Check if user is connecting over VPN
If IsCiscoVPNConnected() Then
	Main.LogFile.Log "Inventory Personal Folder: Identified as VPN connected", True
	Main.LogFile.Log "Inventory Personal Folder: VPN connected Laptop is requested to skip this process", True
	Main.LogFile.Log "Inventory Personal Folder: Completed.", True
' Check if user is part of Retail OU
ElseIf IsRetailUser <> 0 Then
	Main.LogFile.Log "Inventory Personal Folder: User belongs to Retail OU, skipping script execution.", True
	Main.LogFile.Log "Inventory Personal Folder: Completed.", True
Else
	Main.LogFile.Log "Inventory Personal Folder: Identified as VPN not connected", True
	Call GatherAndInsertPersonalFolderInventory(Main)
End If

'If IsOfflineLaptopInventory("Laptop Offline PC") Then
'	Main.LogFile.Log "Inventory Personal Folder: Identified as an Offline Laptop", True
'	Main.LogFile.Log "Inventory Personal Folder: Offline Laptop is requested to skip this process", True
'	Main.LogFile.Log "Inventory Personal Folder: Completed.", True
'	'Offline Laptop should have separate mapping from other desktop types requested by JP IBD'
' Check if user is part of Retail OU
'ElseIf IsRetailUser <> 0 Then
'	Main.LogFile.Log "Inventory Personal Folder: User belongs to Retail OU, skipping script execution.", True
'	Main.LogFile.Log "Inventory Personal Folder: Completed.", True
'Else
'	Call GatherAndInsertPersonalFolderInventory(Main)
'End If

Class InventoryPersonalFolder
	Dim Id
	Dim UserId
	Dim HostName
	Dim Domain
	Dim SiteName
	Dim City
	Dim Path
	Dim UncPath
	Dim Size
	Dim PstLastUpdate
	Dim LastUpdate
	Dim OuMapping

	
	Sub Class_Initialize()
		Id = vbNull
		UserId = vbNull
		HostName = vbNull
		Domain = vbNull
		SiteName = vbNull
		City = vbNull
		Path = vbNull
		UncPath = vbNull
		Size = vbNull
		PstLastUpdate = ""
		LastUpdate = vbNull
		OuMapping = vbNull
	End Sub
	
	
	Function SetupFromXml(Xml)
		Dim xmlNode
		Dim xmlNodes
		
		Set xmlNodes = Xml.SelectNodes("//InventoryPersonalFolder/*")
		
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

				Case "PATH"
					Path = xmlNode.Text

				Case "UNCPATH"
					UncPath = xmlNode.Text
					
				Case "SIZE"
					Size = xmlNode.Text

				Case "PSTLASTUPDATE"
					PstLastUpdate = xmlNode.Text

				Case "LASTUPDATE"
					LastUpdate = xmlNode.Text

				Case "OUMAPPING"
					OuMapping = xmlNode.Text
				End Select
		Next
	End Function
	
	
	Function ToString
		ToString = _
			"INVENTORY PERSONAL FOLDER:" & vbCrLf & _
			"Id: " & Id & vbCrLf &_
			"UserId: " & UserId & vbCrLf &_
			"HostName: " & HostName & vbCrLf &_
			"Domain: " & Domain & vbCrLf &_
			"SiteName: " & SiteName & vbCrLf &_
			"City: " & City & vbCrLf &_
			"Path: " & Path & vbCrLf &_
			"UncPath: " & UncPath & vbCrLf &_
			"Size: " & Size & vbCrLf &_
			"PstLastUpdate: " & PstLastUpdate & vbCrLf &_
			"LastUpdate: " & LastUpdate &_
			"OuMapping: " & OuMapping
	End Function


	Function ToXml
		ToXml = _
			"<InventoryPersonalFolder>" & _
			"<UserId>" & UserId & "</UserId>" & _
			"<HostName>" & HostName & "</HostName>" & _
			"<Domain>" & Domain & "</Domain>" & _
			"<SiteName>" & SiteName & "</SiteName>" & _
			"<City>" & City & "</City>" & _
			"<Path>" & Path & "</Path>" & _
			"<UncPath>" & UncPath & "</UncPath>" & _
			"<Size>" & Size & "</Size>" & _
			"<PstLastUpdate>" & PstLastUpdate & "</PstLastUpdate>" & _
			"<OuMapping>" & OuMapping & "</OuMapping>" & _
			"</InventoryPersonalFolder>"
	End Function
End Class


Class InventoryPersonalFolders
	Dim PersonalFolders()
	Dim PersonalFolderCount

	
	Sub Class_Initialize()
		PersonalFolderCount = 0
	End Sub


	Function ToString
		Dim retVal
		Dim personalFolder
		
		retVal = "INVENTORY PERSONAL FOLDERS:"
		
		For Each personalFolder In PersonalFolders
			retVal = retVal & vbCrLf & Replace(personalFolder.ToString, vbCrLf, vbCrLf & vbTab)
		Next
		
		ToString = retVal
	End Function


	Function ToXml
		Dim retVal
		Dim personalFolder
		
		retVal = "<Mappings>"
		
		For Each personalFolder In PersonalFolders
			retVal = retVal & vbCrLf & Replace(personalFolder.ToXml, vbCrLf, vbCrLf & vbTab)
		Next
		
		ToXml = retVal & "</Mappings>"
	End Function


	Function AddPersonalFolder(UserId, HostName, Domain, OuMapping, SiteName, City, Path, UncPath, Size, PstLastUpdate)
		Dim invPersonalFolder

		Set invPersonalFolder = New InventoryPersonalFolder

		invPersonalFolder.UserId = UserId
		invPersonalFolder.HostName = HostName
		invPersonalFolder.Domain = Domain
		invPersonalFolder.SiteName = SiteName
		invPersonalFolder.City = City
		invPersonalFolder.Path = Path
		invPersonalFolder.UncPath = UncPath
		invPersonalFolder.Size = Size
		invPersonalFolder.PstLastUpdate = PstLastUpdate
		invPersonalFolder.OuMapping = OuMapping

		If (PersonalFolderCount = 0) Then
			ReDim PersonalFolders(0)
		Else
			ReDim Preserve PersonalFolders(UBound(PersonalFolders) + 1)
		End If

		Set PersonalFolders(PersonalFolderCount) = invPersonalFolder

		PersonalFolderCount = PersonalFolderCount + 1
	End Function
End Class


Function InsertPersonalFolderInventory(ByRef Main, PersonalFolders)
	Dim inventoryServer
	Dim strRequest
	Dim http
	Dim intTimeOut

	Main.LogFile.Log "Inventory Personal Folder: About to insert personal folder mapping data to the service", False
	
	intTimeOut = Trim(getResource("InventoryServiceTimeout"))
	
	Main.LogFile.Log "Inventory Personal Folder: Timeout value set as '" & intTimeOut & "' ms", False

	strRequest = ""

	Set inventoryServer = GetInventoryServer(Main.Computer.Domain)

	If (inventoryServer Is Nothing) Then
		Main.LogFile.Log "Inventory Personal Folder: No available inventory service was found", True
		Main.LogFile.Log "Inventory Personal Folder: Completed", True

		Exit Function
	End If

	Main.LogFile.Log "Inventory Personal Folder: Best available inventory service is '" & inventoryServer.ServiceURL & "'", True
	Main.LogFile.Log "Inventory Personal Folder: Server Selection" & vbCrLf & vbCrLf & inventoryServer.ToString & vbCrLf, False

	strRequest = _
		"<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" &_
		" <soap:Body>" &_
		"  <InsertActivePersonalFolderMappingsFromInventory xmlns=""http://webtools.japan.nom"">" &_
		PersonalFolders.ToXml &_
		"  </InsertActivePersonalFolderMappingsFromInventory>" &_
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
			Main.LogFile.Log "Inventory Personal Folder: Exceeded timeout for web service", True
		End If
	Loop
	
	If (Err.Number = 0) Then
		If (intTimeOut > 0) Then
			If (http.Status = 200) Then
				Main.LogFile.Log "Inventory Personal Folders: Received response from web service " & vbCrLf & vbCrLf & http.ResponseText & vbCrLf, False
			Else
				Main.LogFile.Log "Inventory Personal Folders: HTTP request status was not 200. Personal Folder inventory could not be inserted (Status: " & http.Status & ", Response: " & http.ResponseText & ")", True
			End If
		End If
	Else
		Main.LogFile.Log "Inventory Personal Folder: Inventory service unavailable or not responding", True

		Exit Function
	End If

	On Error Goto 0
End Function


Function GatherAndInsertPersonalFolderInventory(ByRef Main)
	Dim mappedPSTs
	Dim InventoryPersonalFolders
	Dim itt

	Main.LogFile.Log "Inventory Personal Folder: Gathering and inserting mapped PSTs", False

	If Not (IfLDAPEmailAttribExists(Main.User.Name) And IsOutlookProfileCreated) Then
		Main.LogFile.Log "Inventory Personal Folder: Outlook profile is not configured for current user.", False

		Exit Function
	End If 

	Set InventoryPersonalFolders = New InventoryPersonalFolders

	mappedPSTs = GetMappedPSTs
	
	On Error Resume Next
		UBound(mappedPSTs)

		If (Err.Number <> 0) Then
			Main.LogFile.Log "Inventory Personal Folder: No PSTs are mapped.", True

			Exit Function
		End If
	On Error GoTo 0

	For itt = 0 To UBound(mappedPSTs) Step 4		
		Main.LogFile.Log "Inventory Personal Folder: Found " & mappedPSTs(itt + 1), True

		If (len(mappedPSTs(itt)) <> 0) Then
			InventoryPersonalFolders.AddPersonalFolder _
				Main.User.Name,_
				Main.Computer.Name,_
				Main.Computer.Domain,_
				Main.User.OuMapping,_
				Main.Computer.Site,_
				Main.Computer.CityCode,_
				mappedPSTs(itt),_
				mappedPSTs(itt + 1),_
				mappedPSTs(itt + 2),_
				mappedPSTs(itt + 3)
		Else
			Main.LogFile.Log "Inventory Personal Folder: Path empty for PST.", False
		End If
	Next

	Main.LogFile.Log "Inventory Personal Folder: Inserting " & InventoryPersonalFolders.PersonalFolderCount & " PSTs", False

	InsertPersonalFolderInventory Main, InventoryPersonalFolders
End Function


Function IsOutlookProfileCreated()
	const HKEY_CURRENT_USER = &H80000001

	Dim RootKey, KeyPath
	Dim arrSubKeys
	Dim oReg
	Dim strComputer

	strComputer = "."
	RootKey = HKEY_CURRENT_USER
	KeyPath = "Software\Microsoft\Office\16.0\Outlook\Profiles"

	Set oReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")
	oReg.EnumKey RootKey, KeyPath, arrSubKeys
	
	If (VarType(arrSubKeys) = vbArray + vbVariant) Then
		isOutlookProfileCreated = True
	Else 
		isOutlookProfileCreated = False 
	End If
End Function 


Function IfLDAPEmailAttribExists(ByVal sUserID)
	Dim objConnection 
	Dim objCommand
	Dim objRecordSet
	Dim objArgs
	Dim strBase
	Dim strFilter
	Dim strQuery
	Dim strAttributes
	Dim OU
	Dim domain
	
	ifLDAPEmailAttribExists = False	

	Select Case Main.Computer.Domain
		Case "ASIAPAC.NOM","QAASIAPAC.NOM","RNDASIAPAC.NOM" 
			OU = "ldap.ap.nomura.com/ou=people,l=ap,o=Nomura.com"

		Case "JAPAN.NOM","QAJAPAN.NOM","RNDJAPAN.NOM"
			OU = "ldap.ap.nomura.com/ou=people,l=ja,o=Nomura.com"

		Case "EUROPE.NOM","QAEUROPE.NOM","RNDEUROPE.NOM"
			OU = "ldap.ap.nomura.com/ou=people,l=eu,o=Nomura.com"

		Case "AMERICAS.NOM","QAAMERICAS.NOM","RNDAMERICAS.NOM"
			OU = "ldap.ap.nomura.com/ou=people,l=us,o=Nomura.com"

		Case Else 
			Main.LogFile.Log "Inventory Personal Folder: Unidentified domain, exiting.", False

			Exit Function 
	End Select 

	Set objConnection = CreateObject("ADODB.Connection")
	objConnection.Provider = "ADsDSOOBject"
	objConnection.Open "Active Directory Provider"

	Set objCommand = CreateObject("ADODB.Command")
	Set objCommand.ActiveConnection = objConnection

	strBase = "<LDAP://" & OU & ">"
	
	'Define the filter elements
	strFilter = "(&(uid=" & sUserID & "))" 

	'List all attributes you will require
	strAttributes = "mail"
	strQuery = strBase & ";" & strFilter & ";" & strAttributes & ";subtree"

	objCommand.CommandText = strQuery

	Set objRecordSet = objCommand.Execute
	
	If (objRecordSet.RecordCount = 0) Then
		ifLDAPEmailAttribExists = False 
	Else
		ifLDAPEmailAttribExists = True 
	End If
End Function 