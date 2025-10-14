Option Explicit


If IsOfflineLaptopMapper("Laptop Offline PC") Then
	Main.LogFile.Log "Mapper Drive: Identified as an Offline Laptop", True
	Main.LogFile.Log "Mapper Drive: Offline Laptop is requested to skip this process", True
	Main.LogFile.Log "Mapper Drive: Completed.", True
	'Offline Laptop should have separate mapping from other desktop types requested by JP IBD'
Else
	Call MapDrives(Main)
End If

Class MapperDrive
	Dim Id
	Dim Domain
	Dim UserId
	Dim AdGroup
	Dim Site
	Dim Drive
	Dim UncPath
	Dim Description
	Dim DisconnectOnLogin
	
	
	Sub Class_Initialize()
		Id = vbNull
		Domain = vbNull
		UserId = vbNull
		AdGroup = vbNull
		Site = vbNull
		Drive = vbNull
		UncPath = vbNull
		Description = ""
		DisconnectOnLogin = vbNull
	End Sub
	
	
	Function SetupFromXml(Xml)
		Dim xmlNode
		Dim xmlNodes
		
		Set xmlNodes = Xml.SelectNodes("//MapperDrive/*")
		
		For Each xmlNode In xmlNodes
			Select Case UCase(xmlNode.baseName)
				Case "ID"
					Id = xmlNode.Text
					
				Case "DOMAIN"
					Domain = xmlNode.Text
					
				Case "USERID"
					UserId = xmlNode.Text
					
				Case "ADGROUP"
					AdGroup = xmlNode.Text
					
				Case "SITE"
					Site = xmlNode.Text
					
				Case "DRIVE"
					Drive = xmlNode.Text
					
				Case "UNCPATH"
					UncPath = xmlNode.Text
					
				Case "DESCRIPTION"
					Description = xmlNode.Text
					
				Case "DISCONNECTONLOGIN"
					DisconnectOnLogin = xmlNode.Text
										
			End Select
		Next
	End Function
	
	
	Function ToString
		ToString = _
			"MAPPER DRIVE:" & vbCrLf & _
			"Id: " & Id & vbCrLf &_
			"Domain: " & Domain & vbCrLf &_
			"UserId: " & UserId & vbCrLf &_
			"AdGroup: " & AdGroup & vbCrLf &_
			"Site: " & Site & vbCrLf &_
			"Drive: " & Drive & vbCrLf &_
			"UncPath: " & UncPath & VbCrLf &_
			"Description: " & Description & VbCrLf &_
			"DisconnectOnLogin: " & DisconnectOnLogin
	End Function
End Class


Class MapperDrives
	Dim Drives()
	
	
	Function SetupFromXml(Xml)
		Dim xmlNode
		Dim xmlNodes
		Dim childXml
		Dim driveLetter
		Dim driveCounter
		Dim insertPosition
		Dim initialLoop
		
		Set xmlNodes = Xml.SelectNodes("//GetUserDrivesResult/*")
		
		initialLoop = True
		
		For Each xmlNode In xmlNodes		
			Set childXml = CreateObject("Microsoft.XMLDOM")
			childXml.async = False
			childXml.LoadXML xmlNode.Xml
			
			driveLetter = childXml.selectSingleNode("//MapperDrive/Drive").Text
			
			'Check and remove conflicts
			If (Not initialLoop) Then
				For driveCounter = 0 To UBound(Drives)
					If (StrComp(driveLetter, Drives(driveCounter).Drive, 1) = 0) Then
						'Update the existing item with new drive mapping (last returned by the service takes precedence)
						insertPosition = driveCounter
						Exit For
					Else
						'A new item will be added to the array
						insertPosition = UBound(Drives) + 1
					End If
				Next
			Else
				ReDim Drives(0)
				insertPosition = 0
			End If
			
			'Increase the size of the array to allow new MapperDrive item to be added
			If (insertPosition > UBound(Drives)) Then
				'Expecting to increase size by 1
				ReDim Preserve Drives(insertPosition)
			End If
			
			Set Drives(insertPosition) = new MapperDrive
			
			Drives(insertPosition).SetupFromXml childXml
			
			initialLoop = False
		Next
	End Function
	
	
	Function ToString
		Dim retVal
		Dim drive
		
		retVal = "MAPPER DRIVES:"
		
		For Each drive In Drives
			retVal = retVal & vbCrLf & Replace(drive.ToString, vbCrLf, vbCrLf & vbTab)
		Next
		
		ToString = retVal
	End Function
End Class


Function GetMapperDrives(ByRef Main)
	Dim mapperServer
	Dim strRequest
	Dim http
	Dim drives
	Dim adGroup
	Dim adGroupsString
	Dim intTimeOut
	
	Main.LogFile.Log "Mapper Drive: About to retrieve drive mapping data from the service", False
	
	intTimeOut = Trim(getResource("MapperServiceTimeout"))
	
	Main.LogFile.Log "Mapper Drive: Timeout value set as '" & intTimeOut & "' ms", False
	
	Set GetMapperDrives = Nothing

	Set mapperServer = GetMapperServer(Main.Computer.Domain)		

	If (mapperServer Is Nothing) Then
		Main.LogFile.Log "Mapper Drive: No available mapper service was found", True
		Main.LogFile.Log "Mapper Drive: Completed", True
		Exit Function
	End If
	
	Main.LogFile.Log "Mapper Drive: Best available mapper service is '" & mapperServer.ServiceURL & "'", True
	Main.LogFile.Log "Mapper Drive: Server Selection" & vbCrLf & vbCrLf & mapperServer.ToString & vbCrLf, False
	
	adGroupsString = ""
	
	If (Not IsNull(Main.User.Groups)) Then
		For Each adGroup in Main.User.Groups
			adGroupsString = adGroupsString &_
				"<string>" & EscapeXMLText(adGroup) & "</string>"
		Next
	End If
	
	strRequest = _
		"<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" &_
		" <soap:Body>" &_
		"  <GetUserDrives xmlns=""http://webtools.japan.nom"">" &_
		"   <UserId xsi:type=""xsd:string"">" & EscapeXMLText(Main.User.Name) & "</UserId>" &_
		"   <Domain xsi:type=""xsd:string"">" & EscapeXMLText(Main.User.ShortDomain) & "</Domain>" &_
		"   <OuMapping xsi:type=""xsd:string"">" & EscapeXMLText(Main.User.OUMapping) & "</OuMapping>" &_
		"   <AdGroups xsi:type=""xsd:string"">" & adGroupsString & "</AdGroups>" &_
		"   <Site xsi:type=""xsd:string"">" & EscapeXMLText(Main.User.CityCode) & "</Site>" &_
		"  </GetUserDrives>" &_
		" </soap:Body>" &_
		"</soap:Envelope>"
		
	Main.LogFile.Log "Mapper Drive: Request content is " & vbCrLf & vbCrLf & strRequest & vbCrLf, False
	
	On Error Resume Next
	
	Set http = CreateObject("Msxml2.XMLHTTP.3.0")
	
	http.Open "POST", mapperServer.ServiceURL, false
	http.SetRequestHeader "Content-Type", "text/xml; CharSet=UTF-8"
	
	http.Send strRequest
	
	'Async request was made. Need to wait until data has all been returned by the server (readystate = 4). The timeout is not 
	Do Until ((http.ReadyState = 4) Or (intTimeOut <= 0))
		intTimeOut = intTimeOut - 100
		Wscript.Sleep 100
		
		If (intTimeOut <= 0) Then
			Main.LogFile.Log "Mapper Drive: Exceeded timeout for web service", True
		End If
	Loop
	
	If (Err.Number = 0) Then
		Main.LogFile.Log "Mapper Drive: HTTP request status is '" & http.Status & "'", True
		
		If (intTimeOut > 0) Then
			If (http.Status = 200) Then
				Main.LogFile.Log "Mapper Drive: Received response from web service " & vbCrLf & vbCrLf & http.ResponseXML.Xml & vbCrLf, False
				Main.LogFile.Log "Mapper Drive: Creating a collection of drive mapping objects", False
				
				Set drives = new MapperDrives
				
				drives.SetupFromXml http.ResponseXML
				
				Set GetMapperDrives = drives
			Else
				Main.LogFile.Log "Mapper Drive: HTTP request status was not 200. Drive mappings could not be retrieved", True
			End If
		End If
	Else
		Main.LogFile.Log "Mapper Drive: Mapper service unavailable or not responding", True
		Exit Function
	End If
	
	On Error Goto 0
End Function


Function MapDrive(ByRef Main, Letter, UncPath, Description, mappedDrives)
	Dim wshNetwork
	Dim wshShell
	Dim shellApplication
	Dim loopCounter
	
	MapDrive = False
	
	Set wshNetwork = WScript.CreateObject("WScript.Network")
	Set wshShell = CreateObject("WScript.Shell")
	Set shellApplication = CreateObject("Shell.Application")
	
	Letter = Letter & ":"
	
	loopCounter = 0
	
	Main.LogFile.Log "Mapper Drive: Expanding environment strings in UncPath '" & UncPath & "'", False
	UncPath = wshShell.ExpandEnvironmentStrings(UncPath)
	Main.LogFile.Log "Mapper Drive: Expanded UncPath to '" & UncPath & "'", False
	
	'Check existing mappings and remove any device names that are already in use and conflict with the device and path that will be mapped
	Do While (loopCounter < UBound(mappedDrives))
			If IsRetailUser <> 0 Then
				If (UCase(Letter) = "V:") Then
					Main.LogFile.Log "Mapper Drive: Skipping V drive mapping for Retail users.", True
					Exit Function
				End If
			End If

		If (UCase(Letter) = UCase(mappedDrives(loopCounter) + ":")) Then
			
			If (UCase(UncPath) = UCase(mappedDrives(loopCounter + 1))) Then
				If IsCiscoVPNConnected() Then
					Main.LogFile.Log "Mapper Drive: Identified as VPN connected", True
					Main.LogFile.Log "Mapper Drive: About to remap '" & Letter & "' to '" & UncPath & "'", True
					wshNetwork.MapNetworkDrive Letter, UncPath, True
				
				else
					Main.LogFile.Log "Mapper Drive: '" & Letter & "' is already mapped to " & UncPath & "'. Will not take any action", True
				end if
				MapDrive = True

				Exit Function
			End If
			
			Main.LogFile.Log "Mapper Drive: '" & Letter & "' is incorrectly mapped to '" & mappedDrives(loopCounter + 1) & "' and will be removed", True
			
			On Error Resume Next
			
			Main.LogFile.Log "Mapper Drive: About to remove mapping", False
			
			wshNetwork.RemoveNetworkDrive Letter, True, True
			
			If (Err.Number = 0) Then
				Main.LogFile.Log "Mapper Drive: Removed incorrectly mapped drive '" & Letter & "'", True
			Else
				Main.LogFile.Log "Mapper Drive: Failed to remove incorrect drive mapping for '" & Letter & "'. " & Err.Description & " (" & Err.Number & ")", True
			End If
			
			On Error Goto 0
			
			Exit Do
		End If
		
		loopCounter = loopCounter + 3
	Loop
	
	'Map the device to the desired path
	Main.LogFile.Log "Mapper Drive: About to map '" & Letter & "' to '" & UncPath & "'", False
	
	On Error Resume Next
	
	wshNetwork.MapNetworkDrive Letter, UncPath, True
	
	If (Err.Number = 0) Then
		MapDrive = True
		Main.LogFile.Log "Mapper Drive: Mapped '" & Letter & "' to '" & UncPath & "'", True
		
		Main.LogFile.Log "Mapper Drive: About to set description for '" & Letter & "' to '" & Description & "'", False
		
		shellApplication.NameSpace(Letter).Self.Name = Description
		
		If (Err.Number = 0) Then
			Main.LogFile.Log "Mapper Drive: Description for '" & Letter & "' set to '" & Description & "'", True
		Else
			Main.LogFile.Log "Mapper Drive: Failed to set description for '" & Letter & "' to '" & Description & "'. " & Err.Description & " (" & Err.Number & ")", True
		End If
	Else
		Main.LogFile.Log "Mapper Drive: Failed to map '" & Letter & "' to '" & UncPath & "'. " & Err.Description & " (" & Err.Number & ")", True
	End If
	
	On Error Goto 0
End Function


Function MapHomeDrive(ByRef Main, mappedDrives)
	Dim wshShell
	Dim myDocumentsPath
	
	Set wshShell = CreateObject("WScript.Shell")
	
	On Error Resume Next
		myDocumentsPath = wshShell.RegRead("HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\Personal")
		
		If (Err.Number = 0) Then
			Main.LogFile.Log "Mapper Drive: About to map home drive to '" & myDocumentsPath & "'", False
			
			MapDrive Main, "H", myDocumentsPath, "Home Drive", mappedDrives
		Else
			Main.LogFile.Log "Mapper Drive: Could not retrieve My Documents location to map home drive to'. " & Err.Description & " (" & Err.Number & ")", True
		End If
	On Error Goto 0
End Function


Function MapDrives(ByRef Main)
	Dim successfulMappings
	Dim failedMappings
	Dim successfulDisconnects
	Dim failedDisconnects
	Dim mappedDrives
	Dim mappedDrive
	Dim mappedDriveCount
	Dim serviceDriveCount
	Dim drives
	Dim drive
	Dim homeDriveMapped
	Dim driveMapped

	Main.LogFile.Log "Mapper Drive: Retrieve drives And paths from service and map in user profile", True
	
	successfulMappings = 0
	failedMappings = 0
	successfulDisconnects = 0
	failedDisconnects = 0
	
	homeDriveMapped = False
	driveMapped = False 
	
	Main.LogFile.Log "Mapper Drive: About to enumerate drives that are already mapped", False
	mappedDrives = GetMappedDrives

	If UBound(mappedDrives) = 0 Then
		Main.LogFile.Log "Mapper Drive: No Drives are mapped.", True
	Else
		Main.LogFile.Log "Mapper Drive: Total of " & ((UBound(mappedDrives)+1) / 3) & " drives already mapped", True			
		driveMapped = True 
	End If

	If driveMapped Then 
		For mappedDriveCount = 0 To UBound(mappedDrives) Step 3
			If (mappedDrives(mappedDriveCount) <> "") Then
				Main.LogFile.Log "Mapper Drive: '" & mappedDrives(mappedDriveCount) & "' is already mapped to '" & mappedDrives(mappedDriveCount + 1) & "'", True
			Else
				Main.LogFile.Log "Mapper Drive: '" & mappedDrives(mappedDriveCount + 1) & "' is mapped with no local device/drive name", True
			End If
			
			If (UCase(mappedDrives(mappedDriveCount) = "H")) Then
				homeDriveMapped = True
			End If
		Next
	End If 

	If IsRetailUser = 0 Then
		'!!! This is required until the environment is sufficiently clean and allows H: to be universally mapped to desired location
		'!!! Only EU and MUM currently allows H: To be mapped To Home drive as per WEAPDMLLS-78 
		If (Main.Computer.ShortDomain = "EUROPE" Or Main.Computer.ShortDomain = "QAEUROPE" OR Main.Computer.ShortDomain = "RNDEUROPE" OR Main.Computer.CityCode = "MUM" _
		OR Main.Computer.ShortDomain = "AMERICAS" Or Main.Computer.ShortDomain = "QAAMERICAS" OR Main.Computer.ShortDomain = "RNDAMERICAS" ) Then
			MapHomeDrive Main, mappedDrives
			homeDriveMapped = True
		Else
			If (homeDriveMapped ) Then
				If IsCiscoVPNConnected() then
				MapHomeDrive Main, mappedDrives
				homeDriveMapped = True
					Main.LogFile.Log "Mapper Drive: VPN connection detected. Remap Home drive letter H:", True

				Else
					Main.LogFile.Log "Mapper Drive: Home drive letter H: is already mapped. Will not (re)map home drive to this letter", True
				End if
			Else
				MapHomeDrive Main, mappedDrives
				homeDriveMapped = True
			End If
		End If
	End If	
	
	Set drives = GetMapperDrives(Main)
	
	If (drives Is Nothing) Then
		Main.LogFile.Log "Mapper Drive: No drives were returned by the service", True
	Else
		Main.LogFile.Log "Mapper Drive: Effective mappings" & vbCrLf & vbCrLf & drives.ToString() & vbCrLf, False
		
		serviceDriveCount = 0
				
		For Each drive in drives.drives
			serviceDriveCount = serviceDriveCount + 1
			
			If (UCase(drive.DisconnectOnLogin) = "TRUE") Then					
				Call DisconnectDrives(Main, drive.drive, drive.UncPath,successfulDisconnects, failedDisconnects)
			End If 
			
			'Refresh the mapped drives array to reflect changes and avoid conflicts
			mappedDrives = GetMappedDrives
				
		Next
				
		For Each drive in drives.drives
						
			If (UCase(drive.DisconnectOnLogin) = "FALSE") Then	
			
				If (MapDrive(Main, drive.drive, drive.UncPath, drive.Description, mappedDrives)) Then
					successfulMappings = successfulMappings + 1
				Else
					failedMappings = failedMappings + 1
				End If
									
			End If
			'Refresh the mapped drives array to reflect changes and avoid conflicts
			mappedDrives = GetMappedDrives
				
		Next
				
		Main.LogFile.Log "Mapper Drive: Total of " & CStr(serviceDriveCount) & " drives returned from service were processed", True
	End If
	
	Main.LogFile.Log "Mapper Drive: " & successfulMappings & " of " & CStr(successfulMappings + failedMappings) & " drives were mapped successfully", True
	Main.LogFile.Log "Mapper Drive: " & successfulDisconnects & " of " & CStr(successfulDisconnects + failedDisconnects) & " drives were disconnected successfully", True
	Main.LogFile.Log "Mapper Drive: Completed", True
End Function


Function DisconnectDrives(ByRef Main, Letter, UncPath,successfulDisconnects, failedDisconnects)
	
	Dim wshNetwork
	Dim wshDrive
	Dim loopCounter
	Dim regExLetter
	Dim regExDrivePath
	Dim letterMatches
	Dim drivePathMatches ' Create variable. 
	Dim driveDisConnectMatchesFound
	Dim mappedDrives
	Dim driveMapped
	Dim mappedDriveCount
     
  	Set wshNetwork = WScript.CreateObject("WScript.Network")
	Set wshDrive = wshNetwork.EnumNetworkDrives
    
	Set regExLetter = New RegExp ' Create a regular expression. 
	Set regExDrivePath = New RegExp 
     
	regExLetter.IgnoreCase = True ' Set case insensitivity. 
	regExLetter.Global = True ' Set global applicability. 
	regExDrivePath.IgnoreCase = True ' Set case insensitivity. 
	regExDrivePath.Global = True ' Set global applicability. 

	Main.LogFile.Log "Mapper Drive: Disconnect requested for Driver Letter" & Letter & ". Drive Path '" & UncPath & "'", False
	'Convert wild card to RegEx 
	uncPath = Replace(uncPath,"\","\\") 
	uncPath = Replace(uncPath,"*",".*") 
	uncPath = Replace(uncPath,"?",".") 
	uncPath = Replace(uncPath,"$","\$") 
	uncPath = "^" & uncPath & "$"
	letter = Replace(letter,"*",".*") 
	letter = Replace(letter,"\","\\") 
	letter = Replace(letter,"?",".") 
    Main.LogFile.Log "Mapper Drive: Disconnect requested for Driver Letter (post wild card to RegEx changes)" & letter & ". Drive Path '" & uncPath & "'", False
	
     
	mappedDrives = GetMappedDrives
	
	DisconnectDrives = False
	driveDisConnectMatchesFound = False
	
	If UBound(mappedDrives) = 0 Then
		Main.LogFile.Log "Mapper Drive: No Drives are mapped.", True
	Else
		Main.LogFile.Log "Mapper Drive: Total of " & ((UBound(mappedDrives)+1) / 3) & " drives already mapped", True			
		driveMapped = True 
	End If

	If driveMapped Then 
		For mappedDriveCount = 0 To UBound(mappedDrives) Step 3
			If (mappedDrives(mappedDriveCount) <> "") Then
				Main.LogFile.Log "Mapper Drive: '" & mappedDrives(mappedDriveCount) & "' is already mapped to '" & mappedDrives(mappedDriveCount + 1) & "'", False

				regExLetter.Pattern = letter ' Set pattern. 
				Set letterMatches = regExLetter.Execute(mappedDrives(mappedDriveCount)) ' Execute search. 
				regExDrivePath.Pattern = uncPath ' Set pattern. 
				Set drivePathMatches = regExDrivePath.Execute(mappedDrives(mappedDriveCount+1)) ' Execute search. 
				
				If letterMatches.Count > 0 And drivePathMatches.Count > 0 Then 
					driveDisConnectMatchesFound = True
					Main.LogFile.Log "Mapper Drive: Found " & mappedDrives(mappedDriveCount) & " " & mappedDrives(mappedDriveCount+1) & " currently mapped", False
					Main.LogFile.Log "Mapper Drive: About to disconnect drive '" & mappedDrives(mappedDriveCount) & "'", False
					
		 			On Error Resume Next
		 	       		wshNetwork.RemoveNetworkDrive mappedDrives(mappedDriveCount) & ":", True, True
						If (Err.Number = 0) Then
							Main.LogFile.Log "Mapper Drive: Successfully disconnected '" & mappedDrives(mappedDriveCount) & "'", True
							DisconnectDrives = True
							successfulDisconnects = successfulDisconnects + 1
						Else
							Main.LogFile.Log "Mapper Drive: Failed to disconnect drive for '" & mappedDrives(mappedDriveCount) & "'. " & Err.Description & " (" & Err.Number & ")", True
							failedDisconnects = failedDisconnects + 1
						End If    		
		    		On Error GoTo 0		
				End If				
			Else
				Main.LogFile.Log "Mapper Drive: '" & mappedDrives(mappedDriveCount + 1) & "' is mapped with no local device/drive name", True
			End If		
		Next
		
		If driveDisConnectMatchesFound = False Then
			failedDisconnects = failedDisconnects + 1
			Main.LogFile.Log "Mapper Drive: None of mapped drive matched the criteria for Disconnecting, doing nothing.", True
		End If 
			
	End If 
End Function