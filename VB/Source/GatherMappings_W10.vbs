 Option Explicit
 
 Const HKEY_CLASSES_ROOT = &H80000000
 Const HKEY_CURRENT_USER = &H80000001
 Const HKEY_LOCAL_MACHINE = &H80000002


Function GetMappedPrinters
	Dim wshNetwork
	
	Set wshNetwork = WScript.CreateObject("WScript.Network")
	
	Set GetMappedPrinters = wshNetwork.EnumPrinterConnections
End Function


Function GetMappedDrives
	Dim rootKey
	Dim key
	Dim subKeys
	Dim subkey
	Dim valueName
	Dim valueData
	Dim regObject
	Dim computerName
	Dim driveMappings()
	Dim firstMapping

	ReDim driveMappings(0)
	firstMapping = True
	
	computerName = "."
	rootKey = HKEY_CURRENT_USER
	key = "Network"
	valueName = "RemotePath"
	
	On Error Resume Next 
	Set regObject=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" &_ 
	   computerName & "\root\default:StdRegProv")
	
	regObject.EnumKey rootKey, key, subKeys
	
	If TypeName(subKeys) <> "Null" Then
		For Each subkey In subKeys
			regObject.GetStringValue rootKey, key & "\" & subkey,valueName,valueData
			If (firstMapping) Then
				firstMapping = False

				ReDim driveMappings(2)
			Else
				ReDim Preserve driveMappings(UBound(driveMappings) + 3)
			End If
			driveMappings(UBound(driveMappings) - 2) = subkey
			driveMappings(UBound(driveMappings) - 1) = valueData
			driveMappings(UBound(driveMappings)) = ""
		Next
	End If 

	On Error GoTo 0
	GetMappedDrives = driveMappings
End Function


Function GetMappedPSTs()
	Const HKEY_CURRENT_USER = &H80000001
	Const masterConfig = "01023d0e"
	Const masterKey = "9207f3e0a3b11019908b08002b2a56c2"
	Const profilesIndex = "Software\Microsoft\Office\16.0\Outlook"
	Const profilesRoot = "Software\Microsoft\Office\16.0\Outlook\Profiles"
	Const defaultProfileString = "DefaultProfile"
	  
	Dim stdRegProv
	Dim masterBinValues
	Dim masterBinValue
	Dim defaultProfileName
	Dim hexNumber
	Dim pstGuid
	Dim pstPath
	Dim pstMappings()
	Dim firstMapping
	
	On Error Resume Next 
		Set stdRegProv = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
		
		firstMapping = True
	
		stdRegProv.GetStringValue HKEY_CURRENT_USER, profilesIndex, defaultProfileString, defaultProfileName	

		If defaultProfileName = Empty Then
			Exit Function 
		End If 
	
		stdRegProv.GetBinaryValue HKEY_CURRENT_USER, profilesRoot & "\" & defaultProfileName & "\" & masterKey, masterConfig, masterBinValues

		For Each masterBinValue In masterBinValues
			If (Len(Hex(masterBinValue)) = 1) Then
				hexNumber = CInt("0") & Hex(masterBinValue)
			Else
				hexNumber = Hex(masterBinValue)
			End If
		
			pstGuid = pstGuid + hexNumber

			If (Len(pstGuid) = 32) Then
				If (IsPST(profilesRoot & "\" & defaultProfileName & "\" & pstGuid)) Then
					pstPath = PstFileName(profilesRoot & "\" & defaultProfileName & "\" & PstLocation(profilesRoot & "\" & defaultProfileName & "\" & pstGuid))

					If (firstMapping) Then
						firstMapping = False

						ReDim pstMappings(3)
					Else
						ReDim Preserve pstMappings(UBound(pstMappings) + 4)
					End If

					pstMappings(UBound(pstMappings) - 3) = pstPath			
					pstMappings(UBound(pstMappings) - 2) = GetUNCPath(pstPath)
					pstMappings(UBound(pstMappings) - 1) = GetFileSize(pstPath)
					pstMappings(UBound(pstMappings)) = GetFileDateLastModified(pstPath)
				End If

				pstGuid = ""
			End If
		Next
	On Error GoTo 0
	
	GetMappedPSTs = pstMappings
End Function


Function IsPST(pstGuid)
	Const HKEY_CURRENT_USER = &H80000001
	Const pstCheck = "00033009"

	Dim pstCheckValues
	Dim pstCheckValue
	Dim stdRegProv
	Dim pstCheckLength
	
	Set stdRegProv = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
	
	pstCheckLength = 0
	IsPST = False
	
	stdRegProv.GetBinaryValue HKEY_CURRENT_USER, pstGuid, pstCheck, pstCheckValues

	For Each pstCheckValue in pstCheckValues
		pstCheckLength = pstCheckLength + Hex(pstCheckValue)
	Next

	If (pstCheckLength = 20) Then
		IsPST = True
	End If
End Function


Function PstLocation(pstGuid)
	Const HKEY_CURRENT_USER = &H80000001
	Const pstGuidLocation = "01023d00"

	Dim pstGuildValues
	Dim pstGuildValue
	Dim stdRegProv
	
	Set stdRegProv = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
	
	stdRegProv.GetBinaryValue HKEY_CURRENT_USER, pstGuid, pstGuidLocation, pstGuildValues

	For Each pstGuildValue In pstGuildValues
		If (Len(Hex(pstGuildValue)) = 1) Then
			PstLocation = pstLocation & CInt("0") & Hex(pstGuildValue)
		Else
			PstLocation = pstLocation & Hex(pstGuildValue)
		End If
	Next
End Function


Function PstFileName(pstGuid)
	Const HKEY_CURRENT_USER = &H80000001
	Const pstFile = "001f6700"

	Dim pstFileValues
	Dim pstFileValue
	Dim tempFileName 
	Dim stdRegProv
	
	set stdRegProv=GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")

	tempFileName = ""
	
	stdRegProv.GetBinaryValue HKEY_CURRENT_USER, pstGuid, pstFile, pstFileValues

	For Each pstFileValue in pstFileValues
		If (pstFileValue > 0) Then
			tempFileName = tempFileName & Chr(pstFileValue)
		End If
	Next

	PstFileName = tempFileName
End Function


Function GetFileSize(FilePath)
	Dim fileSystemObject
	Dim file
	
	Set fileSystemObject = WScript.CreateObject("Scripting.FileSystemObject")

	Set file = fileSystemObject.GetFile(FilePath)

	GetFileSize = file.Size

	Set fileSystemObject = Nothing 
End Function 


Function GetFileDateLastModified(FilePath)
	Dim fileSystemObject
	Dim file
	
	Set fileSystemObject = WScript.CreateObject("Scripting.FileSystemObject")

	Set file = fileSystemObject.GetFile(FilePath)

	GetFileDateLastModified = _
		Year(file.DateLastModified) & "-" &_
		Right("00" & Month(file.DateLastModified), 2) & "-" &_
		Right("00" & Day(file.DateLastModified), 2) & "T" &_
		Right("00" & Hour(file.DateLastModified), 2) & ":" &_
		Right("00" & Minute(file.DateLastModified), 2) & ":" &_
		Right("00" & Second(file.DateLastModified), 2)

	Set fileSystemObject = Nothing 
End Function 


Function GetUNCPath(FilePath)
	Dim itt
	Dim fullUncPath
	Dim networkDrives
	
	fullUncPath = ""
	
	networkDrives = GetMappedDrives
	
	For itt = 0 To UBound(networkDrives) Step 3
		If (StrComp(Left(FilePath,1), networkDrives(itt),vbtextcompare) = 0) Then
			fullUncPath = Replace(FilePath, networkDrives(itt) & ":", networkDrives(itt + 1), 1, 1, vbtextCompare)
		End If 
	Next
	
	If (fullUncPath = "") Then
		GetUNCPath = FilePath
	Else
		GetUNCPath = fullUncPath
	End If 
End Function 