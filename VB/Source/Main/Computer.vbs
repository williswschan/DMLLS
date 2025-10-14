' ===================================================================================
' File:		Users.vbs
'
' Purpose:	Create vbscript User object class for DMLLS scripting framework
'
' Usage:	Computer.vbs
'
' Author:	Calvin Chen
' 
' Updated    	31/05/2024 by Calvin Chen v2.1 - Post CBest Get country code OU logic standardization
'	  				   
' ===================================================================================
Option Explicit

Class ComputerObject
	'Computer's hostname
	Private m_strName
	'Active Directory Distinguished Name
	Private m_strDN
	'Active Directory Domain name
	Private m_DomainDNSName
	'Active Directory short name of the domain
	Private m_DomainShortName
	'Active Directory site of the computer
	Private m_strSite
	'Active Directory group object attributes to query
	Private m_strLdapGroupAttributes
	'Group memberships
	Private m_arrGroups
	'City location code that the computer is based in
	Private m_strCityCode
	'IP addresses
	Private m_arrIPAddress
	'Operating System caption string to identify the OS type
	Private m_strOSCaption
	'Inidcates if Operating System type is desktop (or server)
	Private m_boolDesktop
	
	Private m_boolvpnconnected
	Private m_strFullOU
	'Active Directory Full OU
	
	
	'Name string property (read only)
	Public Property Get Name
		Name = m_strName
	End Property
	
	'DN string property (read only)
	Public Property Get DN
		DN = m_strDN
	End Property
	
	'Domain string property (read only)
	Public Property Get Domain
		Domain = m_DomainDNSName
	End Property
	
	'ShortDomain string property (read only)
	Public Property Get ShortDomain
		ShortDomain = m_DomainShortName
	End Property
		
	'Site string property (read only)
	Public Property Get Site
		Site = m_strSite
	End Property
	
	'GroupAttributes string property (read only)
	Public Property Get GroupAttributes
		GroupAttributes = m_strLdapGroupAttributes
	End Property
	
	'Groups two dimensional array property (read only)
	Public Property Get Groups()
		Groups = m_arrGroups
	End Property
	
	'CityCode string property (read only)
	Public Property Get CityCode
		CityCode = m_strCityCode
	End Property
	
	'IPAddresses array property (read only)
	Public Property Get IPAddresses()
		IPAddresses = m_arrIPAddress
	End Property
	
	'OSCaption string property (read only)
	Public Property Get OSCaption
		OSCaption = m_strOSCaption
	End Property
	
	'Desktop boolean property (read only)
	Public Property Get Desktop
		Desktop = m_boolDesktop
	End Property
	
	

	Public Property Get vpnconnected
		vpnconnected = m_boolvpnconnected
	End Property

	'OU string property (read only)
	Public Property Get OUMapping
		OUMapping = m_strFullOU
	End Property
	
	
	'Constructor. Set the default values
	Private Sub Class_Initialize()
		m_strName = GetName()
		
		Call GetADSystemInfo(m_strDN, m_DomainDNSName, m_DomainShortName, m_strSite)
		m_strLdapGroupAttributes = "distinguishedName,name"
		m_arrGroups = GetGroups(m_strDN, m_strLdapGroupAttributes)
		
		m_strCityCode = GetCityCode(m_strDN)
		
		m_strFullOU = GetOU(m_strDN)
		
		m_arrIPAddress = GetIPAddresses()
		
		m_strOSCaption = GetOSCaption()
		m_boolDesktop = IsDesktopOS()

		m_boolvpnconnected = IsVpnconnected()
		
	End Sub
	
	Private Function GetName()
		Dim objWshNetwork
		
		Set objWshNetwork = WScript.CreateObject("WScript.Network")
		
		GetName = objWshNetwork.ComputerName
	End Function
	
	Private Function GetADSystemInfo(ByRef m_strDN, ByRef m_DomainDNSName, ByRef m_DomainShortName, ByRef m_strSite)
		Dim objADSysInfo
		
		Set objADSysInfo = CreateObject("ADSystemInfo")
		
		m_strDN 			= objADSysInfo.ComputerName
		m_DomainDNSName 	= objADSysInfo.DomainDNSName
		m_DomainShortName 	= objADSysInfo.DomainShortName
		m_strSite 			= objADSysInfo.SiteName
	End Function
	
	'Returns two dimmensional array of desired group attributes (a row is created for each attribute)
	Private Function GetGroups(strDistinguishedName, strLdapGroupAttributes)		
		Dim strDomain, strLdapFilter, strLdapQuery, strAttributeName
		Dim arrForrestDomains, arrLdapGroupAttributes
		Dim objAdSysInfo, objCombinedRecordSet, objConnection, objCommand, objRecordSet
		
		arrLdapGroupAttributes = Split(strLdapGroupAttributes, ",")
		
		On Error Resume Next
		
		'Default return value
		GetGroups = Null
		
		Set objAdSysInfo = CreateObject("ADSystemInfo")
		arrForrestDomains = objAdSysInfo.GetTrees()
		
		'Queries against all domains will be combined and stored in this record set
		Set objCombinedRecordSet = CreateObject("ADODB.Recordset")
		objCombinedRecordSet.cursorLocation = 3 'adUseClient 
		
		For Each strAttributeName in arrLdapGroupAttributes
			objCombinedRecordSet.Fields.Append strAttributeName, 12 'adVariant DataTypeEnum 
		Next
		
		objCombinedRecordSet.Open
		
		If (Err.Number <> 0) Then
			'The recordset that stores group details could not be created
			Exit Function
		End If
		
		For Each strDomain in arrForrestDomains
			'Create query to search for all groups that contain member with supplied distinguished name
			strLdapFilter = "(&(objectClass=group)(member=" & strDistinguishedName & "))"
			strLdapQuery = "<LDAP://" & strDomain & ">;" & strLdapFilter & ";" & strLdapGroupAttributes & ";subtree"
			
			'Create a ADODB.Connection connectin object using ADSI OLE DB
			Set objConnection = CreateObject("ADODB.Connection")
			objConnection.Open "Provider=ADsDSOObject;"
			
			'Create a ADODB.Command object to execute connection with desired query
			Set objCommand = CreateObject("ADODB.Command")
			
			'Assign the connection object and query to command
			With objCommand
				.ActiveConnection = objConnection
				.CommandText = strLdapQuery
			End With
			
			'Execute and store results in record set object
			Set objRecordSet = objCommand.Execute
			
			If (Err.Number = 0) Then
				'The record state may be closed if the command execution did not succeed, check that it is open (state = 1)
				If (objRecordSet.State = 1) Then
					If (objRecordSet.RecordCount > 0) Then
						Do Until objRecordSet.EOF
							'Add values for current record set into combined record set
							objCombinedRecordSet.AddNew
							
							For Each strAttributeName in arrLdapGroupAttributes
								objCombinedRecordSet(strAttributeName) = objRecordSet.Fields.Item(strAttributeName).Value
							Next
							
							'Update the changes and MoveFirst to allow next AddNew
							objCombinedRecordSet.Update
							objCombinedRecordSet.MoveFirst
							
							objRecordSet.MoveNext
						Loop
					End If
				End if
			Else
				Err.Clear
			End If
			
			objConnection.Close
		Next
		
		GetGroups = objCombinedRecordSet.GetRows
		
		objCombinedRecordSet.Close
		
		On Error Goto 0
	End Function
	
	'!!! This method should be updated or replaced when there is a more reliable mechanism for location or the components that require the CityCode property are updated or removed
	'Exctract a city code from the distinguished name. Assumes a structure that includes an OU with this code in the name under a "Resources" OU
	Private Function GetCityCode(strDistinguishedName)
		Dim arrDNParts
		Dim strCityCodeOU
		Dim i
		
		'Default return value
		strCityCodeOU = "unknown"

		If (Instr(1, UCase(strDistinguishedName), "OU=DEVICES,OU=RESOURCESUAT,OU=", 1) > 0) Then
			arrDNParts = Split(strDistinguishedName, ",")
			For i = 0 to Ubound(arrDNParts)
				If (Instr(1, UCase(arrDNParts(i)), "OU=RESOURCESUAT", 1) > 0) Then
				
					strCityCodeOU = UCase(arrDNParts(i +1))
					Exit For
				End If
			Next
		ElseIf (Instr(1, UCase(strDistinguishedName), "OU=DEVICES,OU=RESOURCES,OU=", 1) > 0) Then
			arrDNParts = Split(strDistinguishedName, ",")
			For i = 0 to Ubound(arrDNParts)
				If (Instr(1, UCase(arrDNParts(i)), "OU=RESOURCES", 1) > 0) Then
				wscript.echo i
					strCityCodeOU = UCase(arrDNParts(i + 1))
					Exit For
				End If
			Next
	
		End If
		
		'Remove the "OU=" part of the string and return
		GetCityCode = Replace(strCityCodeOU, "OU=", "", 1, 1, 1)
	End Function
	
	Private Function GetIPAddresses()
		Dim arrIPAddresses
		Dim objWMIService, objNetworkAdapterConfig
		Dim colNetworkAdapterConfigs
		Dim strIPAddress
		
		arrIPAddresses = Array(null)
		
		Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
		Set colNetworkAdapterConfigs = objWMIService.ExecQuery("Select IPAddress from Win32_NetworkAdapterConfiguration where IPEnabled=TRUE")
		
		'Each object in the collection will have an IPAddress property storing the values of the ip addresses as an Array
		For Each objNetworkAdapterConfig in colNetworkAdapterConfigs
			For Each strIPAddress in objNetworkAdapterConfig.IPAddress
				If IsNull(arrIPAddresses(0)) Then
					arrIPAddresses(0) = strIPAddress
				Else
					ReDim Preserve arrIPAddresses(UBound(arrIPAddresses) + 1)
					arrIPAddresses(UBound(arrIPAddresses)) = strIPAddress
				End If
			Next
		Next
		
		GetIPAddresses = arrIPAddresses
	End Function
	
	Function GetOSCaption()
		Dim colComputerSystem
		Dim objComputerSystemItem
		
		On Error Resume Next
		
		Set colComputerSystem = GetObject("winmgmts:\\.\root\cimv2").ExecQuery("Select Caption from Win32_OperatingSystem",,48)
		
		If (Err.Number = 0) Then
			For Each objComputerSystemItem In colComputerSystem		
				GetOSCaption = Trim(objComputerSystemItem.Caption)
			Next
		Else
			GetOSCaption = Empty
		End If
		
		On Error Goto 0
	End Function
	Private Function IsDesktopOS()
		Dim colComputerSystem
		Dim objComputerSystemItem
		Dim intOSrole
		On Error Resume Next
		
		Set colComputerSystem=GetObject("winmgmts:\\.\root\cimv2").ExecQuery("Select DomainRole from Win32_ComputerSystem",,48)
		If (Err.Number = 0) Then
			For Each objComputerSystemItem In colComputerSystem
				intOSRole=objComputerSystemItem.DomainRole
			Next
		Else
			IsDesktopOS = False
		End if
		
		Select Case intOSRole
			Case 0
				IsDesktopOS = True
			Case 1
				IsDesktopOS = True
			Case Else
				IsDesktopOS = False
	
		End Select
		
	End Function	



	Private Function IsVpnconnected()
	
	Dim objWMIService, ObjItem
	Dim colItems
	'On Error Resume Next
	

	Set objWMIService =  GetObject("winmgmts:\\.\root\cimv2")
	Set colItems = objWMIService.ExecQuery ("Select * from Win32_NetworkAdapter where name like 'Cisco AnyConnect%'",,48)
	IsVpnconnected=False
		For Each objItem in colItems

		If objItem.MACAddress <> "" and Instr(Ucase(objItem.Name),"CISCO ANYCONNECT") > 0 Then

	  		if objItem.Netconnectionstatus = "2" then
		
		     		IsVpnconnected=True				
     				
  			end if 
   
		End IF
		Next
	End Function
	
	Public Function ToString()
		Dim strArrItem
		
		ToString =	"Name:           " & vbTab & Me.Name & vbCrLf & _
					"DN:             " & vbTab & Me.DN & vbCrLf & _
					"Domain:         " & vbTab & Me.Domain & vbCrLf & _
					"ShortDomain:    " & vbTab & Me.ShortDomain & vbCrLf & _
					"Site:           " & vbTab & Me.Site & vbCrLf & _
					"OU:             " & vbTab & Me.OUMapping & vbCrLf & _
					"GroupAttributes:" & vbTab & Me.GroupAttributes & vbCrLf
		
		If (IsNull(Me.Groups())) Then
			ToString =	ToString & _
					"Groups():       " & vbTab & "[null]" & vbCrLf
		Else
			For Each strArrItem in Me.Groups()
				ToString =	ToString & _
					"Groups():       " & vbTab & strArrItem & vbCrLf
			Next
		End If
		
		ToString =	ToString & _
					"CityCode:       " & vbTab & Me.CityCode & vbCrLf
		
		If (IsNull(Me.IPAddresses())) Then
			ToString =	ToString & _
					"IPAddresses():  " & vbTab & "[null]" & vbCrLf
		Else
			For Each strArrItem in Me.IPAddresses()
				ToString =	ToString & _
					"IPAddresses():  " & vbTab & strArrItem & vbCrLf
			Next
		End If
		
		ToString =	ToString & _
					"OSCaption:      " & vbTab & Me.OSCaption & vbCrLf & _
					"DesktopOS:        " & vbTab & cStr(Me.Desktop) & vbCrLf & _
					"Vpnconnected:        " & vbTab & cStr(Me.vpnconnected) & vbCrLf
	End Function
	
	Private Function GetOU(strDn)
		Dim strCanonicalName, strDc, strOu, strCn, strDnPart
		Dim arrDnParts
		
		strCanonicalName = ""
		
		'Remove escaped commas and replace with an impossible character combination to avoid splitting DN incorrectly. Convert back later.
		strDn = Replace(strDn, "\,", "##")
		
		arrDnParts = Split(strDn, ",")
		
		For Each strDnPart In arrDnParts
			Select Case Left(UCase(strDnPart), 3)
				Case "DC="
					If (IsEmpty(strDc)) Then
						strDc = Replace(strDnPart, "DC=", "") 
					Else
						strDc = strDc & "." & Replace(strDnPart, "DC=", "")
					End If
				Case "OU="
					If (IsEmpty(strOu)) Then
						strOu = Replace(strDnPart, "OU=", "")
					Else
						strOu = Replace(strDnPart, "OU=", "") & "/" & strOu
					End If
				Case "CN="
					strCn = "/" & Replace(strDnPart, "CN=", "")
			End Select
		Next
		
		'Add preceeding "/" seperator if there are OUs in the distinguished name
		If (Not IsEmpty(strOu)) Then
			strOu = "/" & strOu
		End If
			
		strCanonicalName = strDc & strOu
		
		'Restore previously removed commas and remove DN escape characters 
		strCanonicalName = Replace(strCanonicalName, "##", ",")
		strCanonicalName = Replace(strCanonicalName, "\#", "#")
		strCanonicalName = Replace(strCanonicalName, "\+", "+")
		strCanonicalName = Replace(strCanonicalName, "\<", "<")
		strCanonicalName = Replace(strCanonicalName, "\>", ">")
		strCanonicalName = Replace(strCanonicalName, "\;", ";")
		strCanonicalName = Replace(strCanonicalName, "\""", """")
		strCanonicalName = Replace(strCanonicalName, "\=", "=")
		strCanonicalName = Replace(strCanonicalName, "\ ", " ")
		strCanonicalName = Replace(strCanonicalName, "\/", "/")
		strCanonicalName = Replace(strCanonicalName, "\\", "\") 'Leave this entry last
		
		GetOU = strCanonicalName

	End Function
End Class
