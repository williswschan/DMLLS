' ===================================================================================
' File:		Users.vbs
'
' Purpose:	Create vbscript User object class for DMLLS scripting framework
'
' Usage:	User.vbs
'
' Author:	Calvin Chen
' 
' Updated    	31/05/2024 by Calvin Chen v2.1 - Post CBest Get country code OU logic standardization
'	  				   
' ===================================================================================
Option Explicit

Class UserObject
	'User ID
	Private m_strName
	'Domain name
	Private m_DomainDNSName
	'Short name of the domain
	Private m_DomainShortName
	'Logon Server
	Private m_LogonServer
	'Active Directory Distinguished Name
	Private m_strDN
	'Active Directory group object attributes to query
	Private m_strLdapGroupAttributes
	'Group memberships
	Private m_arrGroups
	'City location code that the user is based in
	Private m_strCityCode
	'Indicates this is a terminal session
	Private m_boolTerminalSession
	Private m_strFullOU
	'Active Directory Full OU
	
	
	'Name string property (read only)
	Public Property Get Name
		Name = m_strName
	End Property
		
	'Domain string property (read only)
	Public Property Get Domain
		Domain = m_DomainDNSName
	End Property
	
	'ShortDomain string property (read only)
	Public Property Get ShortDomain
		ShortDomain = m_DomainShortName
	End Property
	
	'LogonServer string property (read only)
	Public Property Get LogonServer
		LogonServer = m_LogonServer
	End Property
	
	'DN string property (read only)
	Public Property Get DN
		DN = m_strDN
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
	
	'TerminalSession boolean property (read only)
	Public Property Get TerminalSession
		TerminalSession = m_boolTerminalSession
	End Property

	'OU string property (read only)
	Public Property Get OUMapping
		OUMapping = m_strFullOU
	End Property
	
	
	'Constructor. Set the default values
	Private Sub Class_Initialize()
		Call GetUserEnvironmentInfo(m_strName, m_DomainDNSName, m_DomainShortName, m_LogonServer)
		
		m_strDN = GetUserDN()
		
		m_strLdapGroupAttributes = "distinguishedName,name"
		m_arrGroups = GetGroups(m_strDN, m_strLdapGroupAttributes)
		m_strCityCode = GetCityCode(m_strDN)

		m_strFullOU = GetOU(m_strDN)
		
		m_boolTerminalSession = IsTerminalSession()
	End Sub
	
	Private Function GetUserEnvironmentInfo(ByRef m_strName, ByRef m_DomainDNSName, ByRef m_DomainShortName, ByRef m_LogonServer)
		Dim objWshShell
		
		Set objWshShell = WScript.CreateObject("WScript.Shell")
		
		m_strName = objWshShell.ExpandEnvironmentStrings("%USERNAME%")
		m_DomainDNSName = objWshShell.ExpandEnvironmentStrings("%USERDNSDOMAIN%")
		m_DomainShortName = objWshShell.ExpandEnvironmentStrings("%USERDOMAIN%")
		m_LogonServer = objWshShell.ExpandEnvironmentStrings("%LOGONSERVER%")
		
		If (m_strName = "%USERNAME%") Then
			m_strName = ""
		End If
		
		If (m_DomainDNSName = "%USERDNSDOMAIN%") Then
			m_DomainDNSName = ""
		End If
		
		If (m_DomainShortName = "%USERDOMAIN%") Then
			m_DomainShortName = ""
		End If
		
		If (m_LogonServer = "%LOGONSERVER%") Then
			m_LogonServer = ""
		End If
	End Function
	
	Private Function GetUserDN()
		Dim objADSysInfo
		
		Set objADSysInfo = CreateObject("ADSystemInfo")
		
		GetUserDN = objADSysInfo.UserName
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
		If (Instr(1, UCase(strDistinguishedName), "OU=USERS,OU=RESOURCESUAT,OU=", 1) > 0) Then
			arrDNParts = Split(strDistinguishedName, ",")
			For i = 0 to Ubound(arrDNParts)
				If (Instr(1, UCase(arrDNParts(i)), "OU=RESOURCESUAT", 1) > 0) Then
				
					strCityCodeOU = UCase(arrDNParts(i +1))
					Exit For
				End If
			Next
		ElseIf (Instr(1, UCase(strDistinguishedName), "OU=USERS,OU=RESOURCES,OU=", 1) > 0) Then
			arrDNParts = Split(strDistinguishedName, ",")
			For i = 0 to Ubound(arrDNParts)
				If (Instr(1, UCase(arrDNParts(i)), "OU=RESOURCES", 1) > 0) Then
				
					strCityCodeOU = UCase(arrDNParts(i + 1))
					Exit For
				End If
			Next
	
		End If
		'Remove the "OU=" part of the string and return
		GetCityCode = Replace(strCityCodeOU, "OU=", "", 1, 1, 1)
		
	End Function
	
	Private Function IsTerminalSession()
		Dim objWshShell
		Dim strSessionName
		
		Set objWshShell = WScript.CreateObject("WScript.Shell")
		
		strSessionName = objWshShell.ExpandEnvironmentStrings("%SESSIONNAME%")
		
		'Determine if session is terminal by environment string
		If ((strSessionName <> "%SESSIONNAME%") And (strSessionName <> "Console")) Then	
			IsTerminalSession = True
		Else
			IsTerminalSession = False
		End If
	End Function
	
	Public Function ToString()
		Dim strArrItem
		
		ToString =	"Name:           " & vbTab & Me.Name & vbCrLf & _
					"DN:             " & vbTab & Me.DN & vbCrLf & _
					"Domain:         " & vbTab & Me.Domain & vbCrLf & _
					"ShortDomain:    " & vbTab & Me.ShortDomain & vbCrLf & _
					"LogonServer:    " & vbTab & Me.LogonServer & vbCrLf & _
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
		
		ToString =	ToString & _
					"TerminalSession:" & vbTab & cStr(Me.TerminalSession) & vbCrLf
	End Function

	Private Function GetOU(strDn)
		Dim strTempDn, strCanonicalName, strDc, strOu, strCn, strDnPart
		Dim arrDnParts
		
		strCanonicalName = ""
		
		'Remove escaped commas and replace with an impossible character combination to avoid splitting DN incorrectly. Convert back later.
		strTempDn = Replace(strDn, "\,", "##")
		
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
