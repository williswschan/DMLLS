'1st Apr 2020 Calvin Chen  - drop .net 3.5 StringBuilder, replace with wscript.quit
'1st Aug 2023 Calvin Chen  - retire RemovePrinter.vbs logic , line 458-459

Option Explicit

If IsOfflineLaptopMapper("Laptop Offline PC") Then
	Main.LogFile.Log "Mapper Printer: Identified as an Offline Laptop", True
	Main.LogFile.Log "Mapper Printer: Offline Laptop is requested to skip this process", True
	Main.LogFile.Log "Mapper Printer: Completed.", True
	'Offline Laptop should have separate mapping from other desktop types requested by JP IBD'
ElseIf IsUserPartOfGroup("Regional Printer Mapping Inclusion",34) Then
	Main.LogFile.Log "Mapper Printer: User is part of the regional printer mapping inclusion group, printer mapping will continue.", True	
	Main.LogFile.Log "Mapper Printer: Completed.", True
	Call checkMachineOS(Main)
ElseIf Main.User.ShortDomain = "JAPAN" _
			Or Main.User.ShortDomain = "QAJAPAN" _
			Or Main.User.ShortDomain = "RNDJAPAN"  Then
	Main.LogFile.Log "Mapper Printer: User belongs to the domain '" & Main.User.ShortDomain & "', the printer mapping functionality has been disabled by default for this domain.", True
	Main.LogFile.Log "Mapper Printer: Completed.", True
ElseIf IsUserPartOfGroup("Regional Printer Mapping Exclusion",34) Then
	Main.LogFile.Log "Mapper Printer: User is part of the regional printer mapping exclusion group, as such printer mapping process will be skipped.", True	
	Main.LogFile.Log "Mapper Printer: Completed.", True
Else
	Call checkMachineOS(Main)
End If

Function checkMachineOS(ByRef Main)

	If InStr(1,Ucase(Main.Computer.OSCaption),"SERVER") > 0 Then
	    Main.LogFile.Log "Mapper Printer: Operating system of the host machine is '" & Main.Computer.OSCaption & "'. Mapper Printer script will not run.", True
	Else 
	    Call MapPrinters(Main)
	End If
	
End Function  

Class MapperPrinter
	Dim Id
	Dim Domain
	Dim HostName
	Dim AdGroup
	Dim Site
	Dim UncPath
	Dim IsDefault
	Dim Description
	Dim DisconnectOnLogin
	
	
	Sub Class_Initialize()
		Id = vbNull
		Domain = vbNull
		HostName = vbNull
		AdGroup = vbNull
		Site = vbNull
		UncPath = vbNull
		IsDefault = vbNull
		Description = ""
	End Sub
	
	
	Function SetupFromXml(Xml)
		Dim xmlNode
		Dim xmlNodes
		
		Set xmlNodes = Xml.SelectNodes("//MapperPrinter/*")
		
		For Each xmlNode In xmlNodes
			Select Case UCase(xmlNode.baseName)
				Case "ID"
					Id = xmlNode.Text
					
				Case "DOMAIN"
					Domain = xmlNode.Text
					
				Case "HOSTNAME"
					HostName = xmlNode.Text
					
				Case "ADGROUP"
					AdGroup = xmlNode.Text
					
				Case "SITE"
					Site = xmlNode.Text
					
				Case "UNCPATH"
					UncPath = xmlNode.Text
					
				Case "ISDEFAULT"
					IsDefault = xmlNode.Text
					
				Case "DESCRIPTION"
					Description = xmlNode.Text
				
				Case "DISCONNECTONLOGIN"
					DisconnectOnLogin = xmlNode.Text
					
			End Select
		Next
	End Function
	
	
	Function ToString
		ToString = _
			"MAPPER PRINTER:" & vbCrLf &_
			"Id: " & Id & vbCrLf &_
			"Domain: " & Domain & vbCrLf &_
			"HostName: " & HostName & vbCrLf &_
			"AdGroup: " & AdGroup & vbCrLf &_
			"Site: " & Site & vbCrLf &_
			"UncPath: " & UncPath & vbCrLf &_
			"IsDefault: " & IsDefault & VbCrLf &_
			"Description: " & Description & VbCrLf &_
			"DisconnectOnLogin: " & DisconnectOnLogin
	End Function
End Class

Class MapperPrinters
	Dim Printers()
	
	
	Function SetupFromXml(Xml)
		Dim xmlNode
		Dim xmlNodes
		Dim childXml
		Dim printerPath
		Dim printerCounter
		Dim insertPosition
		Dim initialLoop
		
		Set xmlNodes = Xml.SelectNodes("//GetUserPrintersResult/*")
		
		initialLoop = True
		
		For Each xmlNode In xmlNodes
			Set childXml = CreateObject("Microsoft.XMLDOM")
			childXml.async = False
			childXml.LoadXML xmlNode.Xml
			
			printerPath = childXml.selectSingleNode("//MapperPrinter/UncPath").Text
			
			'Check and remove conflicts
			If (Not initialLoop) Then
				For printerCounter = 0 To UBound(Printers)
					If (StrComp(printerPath, Printers(printerCounter).UncPath, 1) = 0) Then
						'Update the existing item with new printer mapping (last returned by the service takes precedence)
						insertPosition = printerCounter
						Exit For
					Else
						'A new item will be added to the array
						insertPosition = UBound(Printers) + 1
					End If
				Next
			Else
				ReDim Printers(0)
				insertPosition = 0
			End If
			
			'Increase the size of the array to allow new MapperPrinter item to be added
			If (insertPosition > UBound(Printers)) Then
				'Expecting to increase size by 1
				ReDim Preserve Printers(insertPosition)
			End If
			
			Set Printers(insertPosition) = new MapperPrinter
			
			Printers(insertPosition).SetupFromXml childXml
			
			initialLoop = False
		Next
	End Function
	
	
	Function ToString
		Dim retVal
		Dim printer
		
		retVal = "MAPPER PRINTERS:"
			
		For Each printer In Printers
			retVal = retVal & vbCrLf & Replace(printer.ToString, vbCrLf, vbCrLf & vbTab)
		Next
		
		ToString = retVal
	End Function
End Class

Function GetMapperPrinters(Main)
	Dim mapperServer
	Dim strRequest
	Dim http
	Dim printers
	Dim adGroup
	Dim adGroupsString
	Dim intTimeOut
	
	Main.LogFile.Log "Mapper Printer: About to retrieve printer mapping data from the service", False
	
	intTimeOut = Trim(getResource("MapperServiceTimeout"))
	
	Main.LogFile.Log "Mapper Printer: Timeout value set as '" & intTimeOut & "' ms", False
	
	Set GetMapperPrinters = Nothing
	Set mapperServer = GetMapperServer(Main.Computer.Domain)
	
	If (mapperServer Is Nothing) Then
		Main.LogFile.Log "Mapper Printer: No available mapper service was found", True
		Main.LogFile.Log "Mapper Printer: Completed", True
		
	        'Set GetMapperPrinters = CreateObject("System.Text.StringBuilder")
	        wscript.quit
		Exit Function
	End If
	
	Main.LogFile.Log "Mapper Printer: Best available mapper service is '" & mapperServer.ServiceURL & "'", True
	Main.LogFile.Log "Mapper Printer: Server Selection" & vbCrLf & vbCrLf & mapperServer.ToString & vbCrLf, False
	
	adGroupsString = ""
	
	If (Not IsNull(Main.Computer.Groups)) Then
		For Each adGroup In Main.Computer.Groups
			adGroupsString = adGroupsString &_
				"<string>" & EscapeXMLText(adGroup) & "</string>"
		Next
	End If
	
	strRequest = _
		"<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" &_
		" <soap:Body>" &_
		"  <GetUserPrinters xmlns=""http://webtools.japan.nom"">" &_
		"   <HostName xsi:type=""xsd:string"">" & EscapeXMLText(Main.Computer.Name) & "</HostName>" &_
		"   <Domain xsi:type=""xsd:string"">" & EscapeXMLText(Main.Computer.ShortDomain) & "</Domain>" &_
		"   <OuMapping xsi:type=""xsd:string"">" & EscapeXMLText(Main.Computer.OUMapping) & "</OuMapping>" &_
		"   <AdGroups xsi:type=""xsd:string"">" & adGroupsString & "</AdGroups>" &_
		"   <Site xsi:type=""xsd:string"">" & EscapeXMLText(Main.Computer.CityCode) & "</Site>" &_
		"  </GetUserPrinters>" &_
		" </soap:Body>" &_
		"</soap:Envelope>"
		
	Main.LogFile.Log "Mapper Printer: Request content is " & vbCrLf & vbCrLf & strRequest & vbCrLf, False
	
	On Error Resume Next
	
	Set http = CreateObject("Msxml2.XMLHTTP.3.0")
	
	http.Open "POST", mapperServer.ServiceURL, True
	http.SetRequestHeader "Content-Type", "text/xml; CharSet=UTF-8"
	
	http.Send strRequest
	
	'Async request was made. Need to wait until data has all been returned by the server (readystate = 4). The timeout is not 
	Do Until ((http.ReadyState = 4) Or (intTimeOut <= 0))
		intTimeOut = intTimeOut - 100
		Wscript.Sleep 100
		
		If (intTimeOut <= 0) Then
			Main.LogFile.Log "Mapper Printer: Exceeded timeout for web service", True
			'Set GetMapperPrinters = CreateObject("System.Text.StringBuilder")
			wscript.quit
		End If
	Loop
	
	If (Err.Number = 0) Then
		Main.LogFile.Log "Mapper Printer: HTTP request status is '" & http.Status & "'", True
		
		If (intTimeOut > 0) Then
			If (http.Status = 200) Then
				Main.LogFile.Log "Mapper Printer: Received response from web service " & vbCrLf & vbCrLf & http.ResponseXML.Xml & vbCrLf, False
				Main.LogFile.Log "Mapper Printer: Creating a collection of printer mapping objects from Web Service.", False
				
				Set printers = new MapperPrinters
				
				printers.SetupFromXml http.ResponseXML
				
				Set GetMapperPrinters = printers
			Else
				Main.LogFile.Log "Mapper Printer: HTTP request status was not 200. Printer mappings could not be retrieved", True
				'Set GetMapperPrinters = CreateObject("System.Text.StringBuilder")
				wscript.quit
			End If
		End If
	Else
		Main.LogFile.Log "Mapper Printer: Mapper service unavailable or not responding", True
		'Set GetMapperPrinters = CreateObject("System.Text.StringBuilder")
		wscript.quit
		Exit Function
	End If
	
	On Error GoTo 0
	
End Function

Function MapPrinter(ByRef Main, UncPath, IsDefault, mappedPrinters)
	Dim wshNetwork
	Dim wshShell
	Dim loopCounter
	Dim isMapped
	
	MapPrinter = False
	
	Set wshNetwork = WScript.CreateObject("WScript.Network")
	Set wshShell = CreateObject("WScript.Shell")
	
	loopCounter = 0
	isMapped = False
	
	Main.LogFile.Log "Mapper Printer: Expanding environment strings in UncPath '" & UncPath & "'", False
	UncPath = wshShell.ExpandEnvironmentStrings(UncPath)
	Main.LogFile.Log "Mapper Printer: Expanded UncPath to '" & UncPath & "'", False
	
	Do While (loopCounter < mappedPrinters.Count - 1)
		If (UCase(UncPath) = UCase(mappedPrinters(loopCounter + 1))) Then
			Main.LogFile.Log "Mapper Printer: '" & UncPath & "' is already mapped. Will not take any action", True
			isMapped = True
			Exit Do
		End If
		
		loopCounter = loopCounter + 2
	Loop
	
	If (isMapped) Then
		'Printer is mapped, return True to indicate success but take no action
		MapPrinter = True
	Else
		'Map the printer
		Main.LogFile.Log "Mapper Printer: About to map printer '" & UncPath & "'", False
		
		On Error Resume Next
		
		wshNetwork.AddWindowsPrinterConnection UncPath
		
		If (Err.Number = 0) Then
			MapPrinter = True
			Main.LogFile.Log "Mapper Printer: Mapped printer '" & UncPath & "'", True
		Else
			Main.LogFile.Log "Mapper Printer: Failed to map printer at path '" & UncPath & "'. " & Err.Description & " (" & Err.Number & ")", True
		End If
		
		On Error Goto 0
	End If
	
	If (IsDefault) Then
		Main.LogFile.Log "Mapper Printer: About to set default printer to: '" & UncPath & "'", False
		
		On Error Resume Next
		
		wshNetwork.SetDefaultPrinter UncPath
		
		If (Err.Number = 0) Then
			Main.LogFile.Log "Mapper Printer: Set default printer as '" & UncPath & "'", True
		Else
			Main.LogFile.Log "Mapper Printer: Failed to set default printer as '" & UncPath & "'. " & Err.Description & " (" & Err.Number & ")", True
		End If
		
		On Error Goto 0
	End If
End Function

Function MapPrinters(ByRef Main)
	Dim successfulMappings
	Dim failedMappings
	Dim successfulDisconnects
	Dim failedDisconnects
	Dim mappedPrinters
	Dim mappedPrinter
	Dim mappedPrinterCount
	Dim servicePrintersCount
	Dim printers
	Dim printer
	Dim mapperPrinterDictionary
	Dim wshShell
	Dim commandline
	Dim strDfsFilePath
	Dim strLocalFilePath
	Dim objFSO
	Dim boolCallScript
	Dim args
			
	Main.LogFile.Log "Mapper Printer: Retrieve printer paths from service and map in user profile", True
		
	successfulMappings = 0
	failedMappings = 0
	successfulDisconnects = 0
	failedDisconnects = 0
	
	Main.LogFile.Log "Mapper Printer: About to enumerate printers that are already mapped", False
	Set mappedPrinters = GetMappedNetworkPrinters
	Main.LogFile.Log "Mapper Printer: Total of " & mappedPrinters.Count & " printers already mapped", True
	
	For Each mappedPrinter In mappedPrinters
		Main.LogFile.Log "Mapper Printer: '" & mappedPrinter.Name & "' is already mapped", True
	Next
	
	Set printers = GetMapperPrinters(Main)
	Set mapperPrinterDictionary = CreateObject("Scripting.Dictionary")
	boolCallScript = False
		
	If (printers Is Nothing) Then
		Main.LogFile.Log "Mapper Printer: No printers were returned by the service", True
		boolCallScript = True
		For Each mappedPrinter In mappedPrinters
			Main.LogFile.Log "Mapper Printer: Disconnecting Printer '" & mappedPrinter.Name & "' as it is not present in GDPMapper database.", True
			Call DisconnectPrinters(Main, mappedPrinter.Name,successfulDisconnects, failedDisconnects)	
		Next
'	ElseIf(TypeName(printers) = "StringBuilder") Then		
'		Main.LogFile.Log "Mapper Printer: Cannot connect to mapper service, will not map/unmap any printers.", True			
	Else
		boolCallScript = True
		Main.LogFile.Log "Mapper Printer: Effective mappings" & VbCrLf & VbCrLf & printers.ToString() & vbCrLf, False
		
		servicePrintersCount = 0
		
		Main.LogFile.Log "Mapper Printer: Comparing Printers returned by the Web Service with the Printers already present on the machine.", True

		For Each printer In printers.printers
			mapperPrinterDictionary.Add UCase(printer.UncPath),""
		Next

		For Each printer In mappedPrinters
			'Check if printer from user's roaming profile is present in GDP Mapper database.
			If mapperPrinterDictionary.Exists(UCase(printer.Name)) <> -1 Then
				Main.LogFile.Log "Mapper Printer: Disconnecting Printer '" & printer.Name & "' as it is not present in GDPMapper database.", True
				Call DisconnectPrinters(Main, printer.Name,successfulDisconnects, failedDisconnects)	
			End If
		Next
			
		For Each printer in printers.printers
			servicePrintersCount = servicePrintersCount + 1
			
			If (UCase(printer.DisconnectOnLogin) = "TRUE") Then	
				Call DisconnectPrinters(Main, printer.UncPath,successfulDisconnects, failedDisconnects)
			End If
					
			'Refresh the mapped printers array to reflect changes and avoid conflicts
			Set mappedPrinters = GetMappedPrinters
		Next
		
		For Each printer in printers.printers
						
			If (UCase(printer.DisconnectOnLogin) = "FALSE") Then	
				If (MapPrinter(Main, printer.UncPath, printer.IsDefault, mappedPrinters)) Then
					successfulMappings = successfulMappings + 1
				Else
					failedMappings = failedMappings + 1
				End If
			End If
					
			'Refresh the mapped printers array to reflect changes and avoid conflicts
			Set mappedPrinters = GetMappedPrinters
		Next		
		
		Main.LogFile.Log "Mapper Printer: Total of " & CStr(servicePrintersCount) & " printers returned from service were processed", True
	End If
	
	If (boolCallScript = True) Then
		Set wshShell = WScript.CreateObject("WScript.Shell")
		Set objFSO = CreateObject("Scripting.FileSystemObject")
		
		'Toggle on off for RemoveUnwantedPrinters.vbs execution
		'strDfsFilePath = "\\" & Main.Computer.Domain & "\Apps\Configfiles\Tools\RemoveUnwantedPrinters.vbs"
		strDfsFilePath =""
		
		If (objFSO.FileExists(strDfsFilePath)) Then
			'Create local copy of file to improve performance when parsing (travelling users)
			strLocalFilePath = wshShell.ExpandEnvironmentStrings("%TEMP%") & "\RemoveUnwantedPrinters.vbs"
			
			objFSO.CopyFile strDfsFilePath, strLocalFilePath, True
			
			If (Err.Number = 0) Then
				Main.LogFile.Log "Mapper Printer: RemoveUnwantedPrinters.vbs copied to '" & strLocalFilePath & "'", False
			Else
				'Use the server path
				strLocalFilePath = strDfsFilePath
			End If

			If (printers Is Nothing) Then
				args = ""
			Else
				For Each printer In printers.printers
					args = args & " """ & printer.UncPath & """"
				Next

				'Remove leading space
				If (Len(args) > 0) Then args = Right(args, Len(args) - 1)
			End If				
						
			commandline = "cscript //NoLogo " & strLocalFilePath & " " & args

			On Error Resume Next
				wshShell.Run commandline, 0, False
			
				If (Err.Number = 0) Then
					Main.LogFile.Log "Mapper Printer: Successfully launched RemoveUnwantedPrinters.vbs script to clean up unwanted printers.", True
				Else
					Main.LogFile.Log "Mapper Printer: Failed to launch RemoveUnwantedPrinters.vbs", True
				End If
			On Error GoTo 0
			
			Main.LogFile.Log "Mapper Printer: " & successfulMappings & " of " & CStr(successfulMappings + failedMappings) & " printers were mapped successfully", True
			Main.LogFile.Log "Mapper Printer: " & successfulDisconnects & " of " & CStr(successfulDisconnects + failedDisconnects) & " printers were disconnected successfully", True		
		Else
			Main.LogFile.Log "Mapper Printer: Cannot access the RemoveUnwantedPrinters.vbs file", True			
		End If		
	End If
	Main.LogFile.Log "Mapper Printer: Completed.", True
End Function

Function DisconnectPrinters(ByRef Main, UncPath,successfulDisconnects, failedDisconnects)

	Dim regExPrinterUncPath
	Dim printerUncPathMatches ' Create variable. 
	Dim printerDisConnectMatchesFound
	Dim mappedPrinters
	Dim printerMapped
	Dim mappedPrintersCount
	Dim currentPrinter
	Dim wshNetwork
	Dim mappedNetworkPrinters
	
	Set wshNetwork = WScript.CreateObject("WScript.Network")	
	Set regExPrinterUncPath = New RegExp ' Create a regular expression. 
	
	regExPrinterUncPath.IgnoreCase = True ' Set case insensitivity. 
	regExPrinterUncPath.Global = True ' Set global applicability. 

	Main.LogFile.Log "Mapper Printer: Disconnect requested for Printer '" & UncPath & "'", False
	'Convert wild card to RegEx 
	uncPath = Replace(uncPath,"\","\\") 
	uncPath = Replace(uncPath,"*",".*") 
	uncPath = Replace(uncPath,"?",".") 
	uncPath = Replace(uncPath,"$","\$") 
	uncPath = "^" & uncPath & "$"
    Main.LogFile.Log "Mapper Printer: Disconnect requested for Printer (post wild card to RegEx changes) '" & uncPath & "'", False
	
	Set mappedPrinters = GetMappedPrinters
	Set mappedNetworkPrinters = GetMappedNetworkPrinters
	
	DisconnectPrinters = False
	printerDisConnectMatchesFound = False
	

	If mappedNetworkPrinters.Count = 0 Then
		Main.LogFile.Log "Mapper Printer: No printers are mapped.", True
	Else
		Main.LogFile.Log "Mapper Printer: Total of " & GetMappedNetworkPrinters.Count & " printers already mapped", True			
		printerMapped = True 
	End If

	If printerMapped Then 
		For mappedPrintersCount = 1 To mappedPrinters.Count step 2
			
			currentPrinter = mappedPrinters.Item(mappedPrintersCount)
			Main.LogFile.Log "Mapper Printer: '" & currentPrinter & "' is already mapped.'", False
		
			regExPrinterUncPath.Pattern = uncPath ' Set pattern. 
			Set printerUncPathMatches = regExPrinterUncPath.Execute(currentPrinter) ' Execute search. 
		
			If printerUncPathMatches.Count > 0 Then 
				printerDisConnectMatchesFound = True
				Main.LogFile.Log "Mapper Printer: Found matching printer for disconnecting - '" & currentPrinter & "'", True
				Main.LogFile.Log "Mapper Printer: About to disconnect Printer '" & currentPrinter & "'", True
			
				On Error Resume Next
		       		wshNetwork.RemovePrinterConnection currentPrinter, True, True
					If (Err.Number = 0) Then
						Main.LogFile.Log "Mapper Printer: Successfully disconnected '" & currentPrinter & "'", True
						DisconnectPrinters = True
						successfulDisconnects = successfulDisconnects + 1
					Else
						Main.LogFile.Log "Mapper Printer: Failed To disconnect printer '" & currentPrinter & "'. " & Err.Description & " (" & Err.Number & ")", True
						failedDisconnects = failedDisconnects + 1
					End If    		
				On Error GoTo 0	
							
			End If				
		Next
	End If 
	
	If printerDisConnectMatchesFound = False Then
		failedDisconnects = failedDisconnects + 1
		Main.LogFile.Log "Mapper Printer: None of mapped printers matched the criteria for Disconnecting, doing nothing.", True
	End If 
	
End Function

Function GetMappedNetworkPrinters
	Dim oWMIService
	
	Set oWMIService = GetObject("winmgmts:" _
 	& "{impersonationLevel=impersonate}!\\" & "." & "\root\cimv2")
 
	Set GetMappedNetworkPrinters = oWMIService.ExecQuery _
 		("Select * from Win32_Printer Where Network = TRUE")
End Function