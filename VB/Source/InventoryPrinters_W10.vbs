Option Explicit

If IsCiscoVPNConnected() Then
	Main.LogFile.Log "Inventory Printer: Identified as VPN connected", True
	Main.LogFile.Log "Inventory Printer: VPN connected Laptop is requested to skip this process", True
	Main.LogFile.Log "Inventory Printer: Completed.", True
	'Offline Laptop should have separate mapping from other desktop types requested by JP IBD'
ElseIf IsUserPartOfGroup("Regional Printer Mapping Exclusion", 34) Then
	Main.LogFile.Log "Inventory Printer: User is part of 'Regional Printer Mapping Exclusion' group, as such printer inventory process will be skipped.", True
	Main.LogFile.Log "Inventory Printer: Completed.", True
	'See GLPE-2118 for details about this request'
' Check if user is part of Retail OU
ElseIf IsRetailUser <> 0 Then
	Main.LogFile.Log "Inventory Printer: User belongs to Retail OU, skipping script execution.", True
	Main.LogFile.Log "Inventory Printer: Completed.", True
Else
	Main.LogFile.Log "Inventory Printer: Identified as VPN not connected", True
	Call GatherAndInsertPrinterInventory(Main)
End If
'If IsOfflineLaptopInventory("Laptop Offline PC") Then
'	Main.LogFile.Log "Inventory Printer: Identified as an Offline Laptop", True
'	Main.LogFile.Log "Inventory Printer: Offline Laptop is requested to skip this process", True
'	Main.LogFile.Log "Inventory Printer: Completed.", True
	'Offline Laptop should have separate mapping from other desktop types requested by JP IBD'
'ElseIf IsUserPartOfGroup("Regional Printer Mapping Exclusion", 34) Then
'	Main.LogFile.Log "Inventory Printer: User is part of 'Regional Printer Mapping Exclusion' group, as such printer inventory process will be skipped.", True
'	Main.LogFile.Log "Inventory Printer: Completed.", True
'	'See GLPE-2118 for details about this request'
' Check if user is part of Retail OU
'ElseIf IsRetailUser <> 0 Then
'	Main.LogFile.Log "Inventory Printer: User belongs to Retail OU, skipping script execution.", True
'	Main.LogFile.Log "Inventory Printer: Completed.", True
'Else
'	Call GatherAndInsertPrinterInventory(Main)
'End If

Class InventoryPrinter
	Dim Id
	Dim UserId
	Dim HostName
	Dim Domain
	Dim SiteName
	Dim City
	Dim UncPath
	Dim IsDefault
	Dim Driver
	Dim Port
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
		UncPath = vbNull
		IsDefault = vbNull
		Driver = vbNull
		Port = vbNull
		Description = ""
		LastUpdate = vbNull
		OuMapping = vbNull
	End Sub
	
	
	Function SetupFromXml(Xml)
		Dim xmlNode
		Dim xmlNodes
		
		Set xmlNodes = Xml.SelectNodes("//InventoryPrinter/*")
		
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

				Case "UNCPATH"
					UncPath = xmlNode.Text
					
				Case "ISDEFAULT"
					IsDefault = xmlNode.Text

				Case "DRIVER"
					Driver = xmlNode.Text

				Case "PORT"
					Port = xmlNode.Text

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
			"INVENTORY PRINTER:" & VbCrLf & _
			"Id: " & Id & vbCrLf &_
			"UserId: " & UserId & vbCrLf &_
			"HostName: " & HostName & vbCrLf &_
			"Domain: " & Domain & vbCrLf &_
			"SiteName: " & SiteName & vbCrLf &_
			"City: " & City & vbCrLf &_
			"UncPath: " & UncPath & vbCrLf &_
			"IsDefault: " & IsDefault & vbCrLf &_
			"Driver: " & Driver & vbCrLf &_
			"Port: " & Port & vbCrLf &_
			"Description: " & Description & vbCrLf &_
			"LastUpdate: " & LastUpdate &_
			"OuMapping: " & OuMapping
	End Function


	Function ToXml
		ToXml = _
			"<InventoryPrinter>" & _
			"<UserId>" & UserId & "</UserId>" & _
			"<HostName>" & HostName & "</HostName>" & _
			"<Domain>" & Domain & "</Domain>" & _
			"<SiteName>" & SiteName & "</SiteName>" & _
			"<City>" & City & "</City>" & _
			"<UncPath>" & UncPath & "</UncPath>" & _
			"<IsDefault>" & IsDefault & "</IsDefault>" & _
			"<Driver>" & Driver & "</Driver>" & _
			"<Port>" & Port & "</Port>" & _
			"<Description>" & Description & "</Description>" & _
			"<OuMapping>" & OuMapping & "</OuMapping>" & _
			"</InventoryPrinter>"
	End Function
End Class


Class InventoryPrinters
	Dim Printers()
	Dim PrinterCount

	
	Sub Class_Initialize()
		PrinterCount = 0
	End Sub


	Function ToString
		Dim retVal
		Dim printer
		
		retVal = "INVENTORY Printers:"
		
		For Each printer In Printers
			retVal = retVal & vbCrLf & Replace(printer.ToString, vbCrLf, vbCrLf & vbTab)
		Next
		
		ToString = retVal
	End Function


	Function ToXml
		Dim retVal
		Dim printer
		
		retVal = "<Mappings>"
		
		For Each printer In Printers
			retVal = retVal & vbCrLf & Replace(printer.ToXml, vbCrLf, vbCrLf & vbTab)
		Next
		
		ToXml = retVal & "</Mappings>"
	End Function


	Function AddPrinter(UserId, HostName, Domain, OuMapping, SiteName, City, UncPath, IsDefault, Driver, Port, Description)
		Dim invPrinter

		Set invPrinter = New InventoryPrinter

		invPrinter.UserId = UserId
		invPrinter.HostName = HostName
		invPrinter.Domain = Domain
		invPrinter.SiteName = SiteName
		invPrinter.City = City
		invPrinter.UncPath = UncPath
		invPrinter.IsDefault = IsDefault
		invPrinter.Driver = Driver
		invPrinter.Port = Port
		invPrinter.Description = Description
		invPrinter.OuMapping = OuMapping

		If (PrinterCount = 0) Then
			ReDim Printers(0)
		Else
			ReDim Preserve Printers(UBound(Printers) + 1)
		End If

		Set Printers(PrinterCount) = invPrinter

		PrinterCount = PrinterCount + 1
	End Function
End Class


Function InsertPrinterInventory(ByRef Main, Printers)
	Dim inventoryServer
	Dim strRequest
	Dim http
	Dim intTimeOut

	Main.LogFile.Log "Inventory Printer: About to insert printer mapping data to the service", False
	
	intTimeOut = Trim(getResource("InventoryServiceTimeout"))
	
	Main.LogFile.Log "Inventory Printer: Timeout value set as '" & intTimeOut & "' ms", False

	strRequest = ""

	Set inventoryServer = GetInventoryServer(Main.Computer.Domain)

	If (inventoryServer Is Nothing) Then
		Main.LogFile.Log "Inventory Printer: No available inventory service was found", True
		Main.LogFile.Log "Inventory Printer: Completed", True

		Exit Function
	End If

	Main.LogFile.Log "Inventory Printer: Best available inventory service is '" & inventoryServer.ServiceURL & "'", True
	Main.LogFile.Log "Inventory Printer: Server Selection" & vbCrLf & vbCrLf & inventoryServer.ToString & vbCrLf, False

	strRequest = _
		"<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" &_
		" <soap:Body>" &_
		"  <InsertMapperPrinterInventory xmlns=""http://webtools.japan.nom"">" &_
		Printers.ToXml &_
		"  </InsertMapperPrinterInventory>" &_
		" </soap:Body>" &_
		"</soap:Envelope>"

	On Error Resume Next

	Main.LogFile.Log "Inventory Printer: Sending request via web service " & vbCrLf & vbCrLf & strRequest & VbCrLf, True
	
	Set http = CreateObject("Msxml2.XMLHTTP.3.0")
	
	http.Open "POST", inventoryServer.ServiceURL, false
	http.SetRequestHeader "Content-Type", "text/xml; CharSet=UTF-8"
		
	http.Send strRequest

	'Async request was made. Need to wait until data has all been returned by the server (readystate = 4). The timeout is not 
	Do Until ((http.ReadyState = 4) Or (intTimeOut <= 0))
		intTimeOut = intTimeOut - 100

		Wscript.Sleep 100
		
		If (intTimeOut <= 0) Then
			Main.LogFile.Log "Inventory Printer: Exceeded timeout for web service", True
		End If
	Loop
	
	If (Err.Number = 0) Then
		If (intTimeOut > 0) Then
			If (http.Status = 200) Then
				Main.LogFile.Log "Inventory Printer: Received response from web service " & vbCrLf & vbCrLf & http.ResponseText & vbCrLf, False
			Else
				Main.LogFile.Log "Inventory Printer: HTTP request status was not 200. Printer inventory could not be inserted (Status: " & http.Status & ", Response: " & http.ResponseText & ")", True
			End If
		End If
	Else
		Main.LogFile.Log "Inventory Printer: Inventory service unavailable or not responding", True

		Exit Function
	End If

	On Error Goto 0
End Function


Function GatherAndInsertPrinterInventory(ByRef Main)
	Dim wmiService
	Dim networkPrinters
	Dim networkPrinter
	Dim inventoryPrinters
	Dim description
	Dim isDefault
	Dim printerShareName
	Dim printerCaption

	Main.LogFile.Log "Inventory Printer: Gathering and inserting printers", False

	Set inventoryPrinters = New InventoryPrinters

	Set wmiService = GetObject("winmgmts:\\.\root\CIMV2")
	Set networkPrinters = wmiService.ExecQuery("SELECT * FROM Win32_Printer")

	For Each networkPrinter In networkPrinters
		If (InStr(networkPrinter.SystemName, "\\") > 0) Then
			If (networkPrinter.Default) Then
				isDefault = 1
			Else
				isDefault = 0
			End If
			
			If (InStr(networkPrinter.ShareName, "&") > 0) Then
				printerShareName = Replace(networkPrinter.ShareName,"&","&amp;")
			Else
				printerShareName = networkPrinter.ShareName
			End If

			
			If (InStr(networkPrinter.Caption, "&") > 0) Then
				printerCaption = Replace(networkPrinter.Caption,"&","&amp;")
			Else
				printerCaption = networkPrinter.Caption
			End If
			inventoryPrinters.AddPrinter _
				Main.User.Name,_
				Main.Computer.Name,_
				Main.Computer.Domain,_
				Main.User.OuMapping,_
				Main.Computer.Site,_
				Main.Computer.CityCode,_
				networkPrinter.SystemName & "\" & printerShareName,_
				isDefault,_
				networkPrinter.DriverName,_
				networkPrinter.PortName,_
				printerCaption
		End If
	Next

	Main.LogFile.Log "Inventory Printer: Inserting " & inventoryPrinters.PrinterCount & " printers", False

	InsertPrinterInventory Main, inventoryPrinters
End Function