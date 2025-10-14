Option Explicit


If IsOfflineLaptopMapper("Laptop Offline PC") or IsCiscoVPNConnected() Then
	Main.LogFile.Log "Mapper PST: Identified as VPN connected or Offline Laptop", True
	Main.LogFile.Log "Mapper PST: Request to skip this process", True
	Main.LogFile.Log "Mapper PST: Skipping Completed.", True
	'Offline Laptop should have separate mapping from other desktop types requested by JP IBD'
ElseIf IsRetailUser <> 0 Then
	Main.LogFile.Log "Mapper PST: User belongs to Retail OU, skipping script execution.", True
	Main.LogFile.Log "Mapper PST: Skipping Completed.", True
Else
	Call checkMachineOS(Main)
End If

Function checkMachineOS(ByRef Main)

	If InStr(1,Ucase(Main.Computer.OSCaption),"SERVER") > 0 Then
	    Main.LogFile.Log "Mapper PST: Operating system of the host machine is '" & Main.Computer.OSCaption & "'. Mapper PST script will not run.", True
	Else 
	    Call MapPersonalFolders(Main)
	End If
End Function  

Class MapperPersonalFolder
	Dim Id
	Dim Domain
	Dim UserId
	Dim AdGroup
	Dim Site
	Dim UncPath
	Dim DisconnectOnLogin
	
	
	Sub Class_Initialize()
		Id = vbNull
		Domain = vbNull
		UserId = vbNull
		AdGroup = vbNull
		Site = vbNull
		UncPath = vbNull
	End Sub
	
	
	Function SetupFromXml(Xml)
		Dim xmlNode
		Dim xmlNodes
		
		Set xmlNodes = Xml.SelectNodes("//MapperPersonalFolder/*")
		
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
					
				Case "UNCPATH"
					UncPath = xmlNode.Text
							
				Case "DISCONNECTONLOGIN"
					DisconnectOnLogin = xmlNode.Text					
			End Select
		Next
	End Function

	
	Function ToString
		ToString = _
			"MAPPER PERSONAL FOLDER:" & vbCrLf &_
			"Id: " & Id & vbCrLf &_
			"Domain: " & Domain & vbCrLf &_
			"UserId: " & UserId & vbCrLf &_
			"AdGroup: " & AdGroup & vbCrLf &_
			"Site: " & Site & vbCrLf &_
			"UncPath: " & UncPath & vbCrLf &_
			"DisconnectOnLogin: " & DisconnectOnLogin
	End Function
End Class


Class MapperPersonalFolders
	Dim PersonalFolders()

	
	Function SetupFromXml(Xml)
		Dim xmlNode
		Dim xmlNodes
		Dim childXml
		Dim pstPath
		Dim pstCounter
		Dim insertPosition
		Dim initialLoop
		
		Set xmlNodes = Xml.SelectNodes("//GetUserPersonalFoldersResult/*")
		
		initialLoop = True
		
		For Each xmlNode In xmlNodes
			Set childXml = CreateObject("Microsoft.XMLDOM")
			childXml.async = False
			childXml.LoadXML xmlNode.Xml
			
			pstPath = childXml.selectSingleNode("//MapperPersonalFolder/UncPath").Text
			
			'Check and remove conflicts
			If (Not initialLoop) Then
				For pstCounter = 0 To UBound(PersonalFolders)
					If (StrComp(pstPath, PersonalFolders(pstCounter).UncPath, 1) = 0) Then
						'Update the existing item with new pst mapping (last returned by the service takes precedence)
						insertPosition = pstCounter
						Exit For
					Else
						'A new item will be added to the array
						insertPosition = UBound(PersonalFolders) + 1
					End If
				Next
			Else
				ReDim PersonalFolders(0)
				insertPosition = 0
			End If
			
			'Increase the size of the array to allow new MapperPersonalFolder item to be added
			If (insertPosition > UBound(PersonalFolders)) Then
				'Expecting to increase size by 1
				ReDim Preserve PersonalFolders(insertPosition)
			End If
			
			Set PersonalFolders(insertPosition) = new MapperPersonalFolder
			
			PersonalFolders(insertPosition).SetupFromXml childXml
			
			initialLoop = False
		Next
	End Function

	
	Function ToString
		Dim retVal
		Dim personalFolder
		
		retVal = "MAPPER PERSONAL FOLDERS:"
			
		For Each personalFolder In PersonalFolders
			retVal = retVal & VbCrLf & Replace(personalFolder.ToString, vbCrLf, vbCrLf & vbTab)
		Next
		
		ToString = retVal
	End Function
End Class


Function GetMapperPersonalFolders(ByRef Main)
	Dim mapperServer
	Dim strRequest
	Dim http
	Dim personalFolders
	Dim intTimeOut
	
	Main.LogFile.Log "Mapper PST: About to retrieve PST mapping data from the service", False
	
	intTimeOut = Trim(getResource("MapperServiceTimeout"))
	
	Main.LogFile.Log "Mapper PST: Timeout value set as '" & intTimeOut & "' ms", False
	
	Set GetMapperPersonalFolders = Nothing
	Set mapperServer = GetMapperServer(Main.Computer.Domain)
	
	If (mapperServer Is Nothing) Then
		Main.LogFile.Log "Mapper PST: No available mapper service was found", True
		Main.LogFile.Log "Mapper PST: Completed", True
		Exit Function
	End If
	
	Main.LogFile.Log "Mapper PST: Best available mapper service is '" & mapperServer.ServiceURL & "'", True
	Main.LogFile.Log "Mapper PST: Server Selection" & vbCrLf & vbCrLf & mapperServer.ToString & vbCrLf, False
	
	strRequest = _
		"<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" &_
		" <soap:Body>" &_
		"  <GetUserPersonalFolders xmlns=""http://webtools.japan.nom"">" &_
		"   <UserId xsi:type=""xsd:string"">" & EscapeXMLText(Main.User.Name) & "</UserId>" &_
		"   <Domain xsi:type=""xsd:string"">" & EscapeXMLText(Main.User.ShortDomain) & "</Domain>" &_
		"   <OuMapping xsi:type=""xsd:string"">" & EscapeXMLText(Main.User.OUMapping) & "</OuMapping>" &_
		"  </GetUserPersonalFolders>" &_
		" </soap:Body>" &_
		"</soap:Envelope>"
	
	Main.LogFile.Log "Mapper PST: Request content is " & vbCrLf & vbCrLf & strRequest & vbCrLf, False
	
	On Error Resume Next
	
	Set http = CreateObject("Msxml2.XMLHTTP.3.0")
	
	http.Open "POST", mapperServer.ServiceURL, False
	http.SetRequestHeader "Content-Type", "text/xml; CharSet=UTF-8"
	
	http.Send strRequest
	
	'Async request was made. Need to wait until data has all been returned by the server (readystate = 4). The timeout is not 
	Do Until ((http.ReadyState = 4) Or (intTimeOut <= 0))
		intTimeOut = intTimeOut - 100
		Wscript.Sleep 100
		
		If (intTimeOut <= 0) Then
			Main.LogFile.Log "Mapper PST: Exceeded timeout for web service", True
		End If
	Loop
	
	If (Err.Number = 0) Then
		Main.LogFile.Log "Mapper PST: HTTP request status is '" & http.Status & "'", True
		
		If (intTimeOut > 0) Then
			If (http.Status = 200) Then
				Main.LogFile.Log "Mapper PST: Received response from web service " & vbCrLf & vbCrLf & http.ResponseXML.Xml & vbCrLf, False
				Main.LogFile.Log "Mapper PST: Creating a collection of personal folder mapping objects", False
				
				Set personalFolders = new MapperPersonalFolders
				
				personalFolders.SetupFromXml http.ResponseXML
				
				Set GetMapperPersonalFolders = personalFolders
			Else
				Main.LogFile.Log "Mapper PST: HTTP request status was not 200. Personal folder mappings could not be retrieved", True
			End If
		End If
	Else
		Main.LogFile.Log "Mapper PST: Mapper service unavailable or not responding", True
		Exit Function
	End If
	
	On Error GoTo 0
End Function


Function MapPersonalFolder(ByRef Main, UncPath)
	Dim fso
	Dim wshShell
	Dim outlook
	Dim nsMapi
	
	MapPersonalFolder = False
	
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set wshShell = CreateObject("WScript.Shell")
	
	Main.LogFile.Log "Mapper PST: Expanding environment strings in UncPath '" & UncPath & "'", False
	UncPath = wshShell.ExpandEnvironmentStrings(UncPath)
	Main.LogFile.Log "Mapper PST: Expanded UncPath to '" & UncPath & "'", False
	
	If (Not fso.FileExists(UncPath)) Then
		Main.LogFile.Log "Mapper PST: Could not find the file '" & UncPath & "'", True
		Exit Function
	End If
	
	On Error Resume Next
	
	Set outlook = CreateObject("Outlook.Application")
	Set nsMapi = outlook.GetNameSpace("MAPI")
	
	Main.LogFile.Log "Mapper PST: About to map PST '" & UncPath & "'", True
	
	nsMapi.AddStore UncPath
	
	If (Err.Number = 0) Then
		MapPersonalFolder = True
		Main.LogFile.Log "Mapper PST: Mapped personal folder '" & UncPath & "'", True
	Else
		Main.LogFile.Log "Mapper PST: Failed to map personal folder at path '" & UncPath & "'. " & Err.Description & " (" & Err.Number & ")", True
	End If
	
	On Error GoTo 0
End Function


Function MapPersonalFolders(ByRef Main)
	Dim successfulMappings
	Dim failedMappings
	Dim successfulDisconnects
	Dim failedDisconnects
	Dim personalFolders
	Dim personalFolder
	Dim servicePSTCount
	
	Main.LogFile.Log "Mapper PST: Retrieve personal folders (PST) locations from service and map in Outlook", True
	
	successfulMappings = 0
	failedMappings = 0
	successfulDisconnects = 0
	failedDisconnects = 0
	
	Set personalFolders = GetMapperPersonalFolders(Main)
	
	If (personalFolders Is Nothing) Then
		Main.LogFile.Log "Mapper PST: No personal folders were returned by the service", True
	Else
		Main.LogFile.Log "Mapper PST: Effective Mappings" & vbCrLf & vbCrLf & personalFolders.ToString() & vbCrLf, False
		
		servicePSTCount = 0
		
		For Each personalFolder In personalFolders.personalFolders
			servicePSTCount = servicePSTCount + 1
			
			If (UCase(personalFolder.DisconnectOnLogin) = "TRUE") Then
				Call DisconnectPersonalFolder(Main, personalFolder.UncPath, successfulDisconnects, failedDisconnects)				
			End If
		Next
		
		For Each personalFolder In personalFolders.personalFolders
						
			If (UCase(personalFolder.DisconnectOnLogin) = "FALSE") Then
			
				If (MapPersonalFolder(Main, personalFolder.UncPath)) Then
					successfulMappings = successfulMappings + 1
				Else
					failedMappings = failedMappings + 1
				End If			
											
			End If
				
		Next	
		
		Main.LogFile.Log "Mapper PST: Total of " & CStr(servicePSTCount) & " personal folders returned from service were processed", True
	End If
	
	Main.LogFile.Log "Mapper PST: " & successfulMappings & " of " & CStr(successfulMappings + failedMappings) & " personal folders were mapped successfully", True
	Main.LogFile.Log "Mapper PST: " & successfulDisconnects & " of " & CStr(successfulDisconnects + failedDisconnects) & " personal folders were disconnected successfully", True
	Main.LogFile.Log "Mapper PST: Completed", True
End Function


Function DisconnectPersonalFolder(ByRef Main, UncPath, successfulDisconnects, failedDisconnects)

	Dim outlook
	Dim nsMapi
	Dim f
	Dim fcount
	Dim regExPSTPath
	Dim PSTPathMatches 
	Dim PSTDisConnectMatchesFound
	Dim mappedPersonalFolders
	Dim PSTMapped
	Dim mappedPSTPath
		
	Set regExPSTPath = New RegExp ' Create a regular expression.
	
	regExPSTPath.IgnoreCase = True ' Set case insensitivity. 
	regExPSTPath.Global = True ' Set global applicability. 

	Main.LogFile.Log "Mapper PST: Disconnect requested for PST Path " & UncPath, False
	'Convert wild card to RegEx 
	uncPath = Replace(uncPath,"\","\\") 
	uncPath = Replace(uncPath,"*",".*") 
	uncPath = Replace(uncPath,"?",".") 
	uncPath = Replace(uncPath,"$","\$") 
	uncPath = "^" & uncPath & "$"
	Main.LogFile.Log "Mapper PST: Disconnect requested for Driver Path (post wild card to RegEx changes): '" & uncPath & "'", False

	DisconnectPersonalFolder = False
	PSTDisConnectMatchesFound = False 
	
	On Error Resume Next

	Set outlook = CreateObject("Outlook.Application") 
   	Set nsMapi = outlook.GetNamespace("MAPI")
   	
	fcount = nsMapi.Folders.count
	For f = fcount To 1 Step -1
		mappedPSTPath = GetPSTPath(nsMapi.Folders(f).StoreID)

		regExPSTPath.Pattern = uncPath ' Set pattern. 
		Set PSTPathMatches = regExPSTPath.Execute(mappedPSTPath) ' Execute search. 
		
		If mappedPSTPath <> "" Then 
			Main.LogFile.Log "Mapper PST: Found " & mappedPSTPath & " " & " is currently mapped", False

			If PSTPathMatches.Count > 0 Then 	
				PSTDisConnectMatchesFound = True			
				Main.LogFile.Log "Mapper PST: About to disconnect PST '" & mappedPSTPath & "'", False
				nsMapi.RemoveStore (nsMapi.Folders(f))
				
				If (Err.Number = 0) Then
					Main.LogFile.Log "Mapper PST: Successfully disconnected '" & mappedPSTPath & "'", True
					DisconnectPersonalFolder = True
					successfulDisconnects = successfulDisconnects + 1
				Else
					Main.LogFile.Log "Mapper PST: Failed to disconnect PST for '" & mappedPSTPath & "'. " & Err.Description & " (" & Err.Number & ")", True
					failedDisconnects = failedDisconnects + 1
				End If    	
		
			End If
		End If 
	Next	
	On Error GoTo 0

	If PSTDisConnectMatchesFound = False Then
		failedDisconnects = failedDisconnects + 1
		Main.LogFile.Log "Mapper PST: None of mapped PST matched the criteria for Disconnecting, doing nothing.", True
	End If 
End Function

Function GetPSTPath(StoreID)
	Dim strPath
	Dim i
	Dim strSubString

    For i = 1 To Len(StoreID) Step 2
        strSubString = Mid(StoreID,i,2)    
        If Not strSubString = "00" Then strPath = strPath & ChrW("&H" & strSubString)
    Next
   
    Select Case True
        Case InStr(strPath,":\") > 0  
            GetPSTPath = Mid(strPath,InStr(strPath,":\")-1)
        Case InStr(strPath,"\\") > 0  
            GetPSTPath = Mid(strPath,InStr(strPath,"\\"))
    End Select
End Function