Option Explicit

'Build contents of a log file and write to specified location
Class LoggingObject
	'The text file path
	Private m_strPath
	'The contents of the text file
	Private m_strContents
	'If verbose logging is enabled, all content will be added to log
	Private m_boolVerboseLogging
	'Last error number
	Private m_lngErrorLevel
	'Last error description
	Private m_strErrorDescription
	'Enables Log to be kept in memory until Write, Append or Clear is called
	Private m_boolCacheLog
	
	'FileSystemObject that will be used by various methods
	Private m_objFSO
	
	
	'Path string property
	Public Property Get Path
		Path = m_strPath
	End Property
	
	Public Property Let Path(strPath)
		m_strPath = strPath
	End Property
	
	'Contents string property (read only)
	Public Property Get Contents
		Contents = m_strContents
	End Property
	
	'VerboseLogging boolean property
	Public Property Get VerboseLogging
		VerboseLogging = m_boolVerboseLogging
	End Property
	
	Public Property Let VerboseLogging(boolVerboseLogging)
		m_boolVerboseLogging = boolVerboseLogging
	End Property
	
	'ErrorLevel long property (read only)
	Public Property Get ErrorLevel
		ErrorLevel = m_lngErrorLevel
	End Property
	
	'ErrorDescription string property (read only)
	Public Property Get ErrorDescription
		ErrorDescription = m_strErrorDescription
	End Property
	
	'CacheLog boolean property
	Public Property Get CacheLog
		CacheLog = m_boolCacheLog
	End Property
	
	Public Property Let CacheLog(boolCacheLog)
		m_boolCacheLog = boolCacheLog
	End Property
	
	
	'Constructor. Set the default values
	Private Sub Class_Initialize()
		Set m_objFSO = CreateObject("Scripting.FileSystemObject")
		
		m_strPath = Empty
		m_strContents = ""
		
		m_boolVerboseLogging = False
		m_boolCacheLog = False
		
		m_lngErrorLevel = 0
		m_strErrorDescription = ""
	End Sub
	
	'Add carriage return and string  with timestamp to Contents. Conditional; will add everything if verbose logging enabled or if boolAlwaysLog
	Public Function Log(strText, boolAlwaysLog)
		On Error Resume Next
		
		'Default return value
		Log = 0
		
		'Only update contents if required
		If (m_boolVerboseLogging Or boolAlwaysLog) Then
			
			'Prevent unnecessary line break at beginning of file
			If (Len(m_strContents) > 0) Then
				m_strContents = m_strContents & vbCrLf
			End If
			
			strText = TimeStamp(Now(), True) & vbTab & strText
			m_strContents = m_strContents & strText
			
			If (Not m_boolCacheLog) Then
				Log = AppendContentToFile
			End If
		End If
		
		If (Err.Number = 0) Then
			'Clear any previous errors as this method executed successfully
			ClearError
		Else
			SetError Err.Number, Err.Description
			Log = Err.Number
			Exit Function
		End If
		
		On Error Goto 0
	End Function
	
	'Update Contents with string
	Public Function Update(strText)
		On Error Resume Next
		
		m_strContents = m_strContents & strText
		
		If (Err.Number <> 0) Then
			SetError Err.Number, Err.Description
			Update = Err.Number
			Exit Function
		End If
		
		On Error Goto 0
		
		'Clear any previous errors as this method executed successfully
		ClearError
		
		Update = 0
	End Function
	
	'Write or Append Contents to Path
	Private Function WriteOrAppendContentToFile(boolNewFile)
		Dim objFile
		Dim intCreateParentFolderReturn
		
		If (IsNull(m_strPath) Or IsEmpty(m_strPath)) Then
			'Do not write anything
			ClearError
			Exit Function
		End If
		
		intCreateParentFolderReturn = CreateParentFolder(m_strPath)
		
		If (intCreateParentFolderReturn <> 0) Then
			WriteOrAppendContentToFile = intCreateParentFolderReturn
			Exit Function
		End If
		
		On Error Resume Next
		
		If (boolNewFile) Then
			'Open for writing
			Set objFile = m_objFSO.OpenTextFile(m_strPath, 2, True, -1)
		Else
			'Open for appending
			Set objFile = m_objFSO.OpenTextFile(m_strPath, 8, True, -1)
		End If
		
		If (Err.Number <> 0) Then
			SetError Err.Number, Err.Description
			WriteOrAppendContentToFile = Err.Number
			Exit Function
		End If
		
		objFile.WriteLine m_strContents
		
		If (Err.Number <> 0) Then
			SetError Err.Number, Err.Description
			WriteOrAppendContentToFile = Err.Number
			Exit Function
		End If
		
		objFile.Close
		Set objFile = Nothing
		
		On Error Goto 0
		
		'Clear any previous errors as this method executed Successfully
		ClearError
		
		'Clear text file contents on successful write
		ClearContents
		
		WriteOrAppendContentToFile = 0
	End Function
	
	'Write Contents to file (overwrite existing file)
	Public Function WriteContentToFile()
		WriteContentToFile = WriteOrAppendContentToFile(True)
	End Function
	
	'Append Contents to file (preserve any existing content)
	Public Function AppendContentToFile()
		Dim objFile
		
		On Error Resume Next
		
		'If the file exists then append to new line for correct formating
		If (m_objFSO.FileExists(m_strPath)) Then
			Set objFile = m_objFSO.OpenTextFile(m_strPath, 1, False, -1)
			
			If (Err.Number <> 0) Then
				SetError Err.Number, Err.Description
				AppendContentToFile = Err.Number
				Exit Function
			End If
			
			'Read all to find column value
			objFile.ReadAll
			
			If (Err.Number <> 0) Then
				SetError Err.Number, Err.Description
				AppendContentToFile = Err.Number
				Exit Function
			End If
			
			'Ensure that any new contents is written to a new line
			If (objFile.Column <> 1) Then
				m_strContents = vbCrLf & m_strContents
			End If
			
			Set objFile = Nothing
			
			On Error Goto 0
		End If
		
		AppendContentToFile = WriteOrAppendContentToFile(False)
	End Function
	
	'Update content and immediately write it to the file
	Public Function Write(strText)
		Update strText
		Write = WriteContentToFile
	End Function
	
	'Update content and immediately append it to the file
	Public Function Append(strText)
		Update strText
		Append = AppendContentToFile
	End Function
	
	'Reset value for m_strContents
	Public Function ClearContents()
		m_strContents = ""
		ClearError
	End Function
	
	'Get the date and time with or without seperators in a consistent format
	Public Function TimeStamp(dateDateTime, boolSeperators)
		Dim strDateTime, strDateSeperator, strTimeSeperator, strDateTimeSeperator
		
		If (boolSeperators) Then
			'"yyyy/MM/dd HH:mm:ss"
			strDateSeperator = "/"
			strTimeSeperator = ":"
			strDateTimeSeperator = " "
		Else
			'"yyyyMMddHHmmss"
			strDateSeperator = ""
			strTimeSeperator = ""
			strDateTimeSeperator = ""
		End If
		
		If (IsDate(dateDateTime)) Then
			strDateTime = Year(dateDateTime) & strDateSeperator & PadNumber(Month(dateDateTime),2) & strDateSeperator & PadNumber(Day(dateDateTime),2)
			strDateTime = strDateTime & strDateTimeSeperator & PadNumber(Hour(dateDateTime),2) & strTimeSeperator & PadNumber(Minute(dateDateTime),2) & strTimeSeperator & PadNumber(Second(dateDateTime),2)
			
			'Clear any previous errors as this method executed Successfully
			ClearError
		Else
			SetError 13, "Type mismatch"
			strDateTime = CStr(dateDateTime)
		End If
		
		TimeStamp = strDateTime
	End Function

	'Recursively pad a string with preceeding zeros until it reaches required length
	Private Function PadNumber(strNumber, intRequiredLength)
		If (Len(strNumber) < intRequiredLength) Then
			PadNumber = PadNumber("0" & strNumber, intRequiredLength)
		Else
			PadNumber = strNumber
		End If
	End Function
	
	'Create parent folder(s) for a given file path
	Private Function CreateParentFolder(strPath)
		Dim arrPathParts
		Dim boolUNCPath
		Dim intCountParts, intServerDriveItems
		Dim strParentFolderPart
		
		If (InStr(strPath, ":\") = 0) Then
			boolUNCPath = True
			
			'The number of items in the array that should be ignored because it is the server name
			intServerDriveItems = 3
		Else
			If (Left(strPath, 4) = "\\?\") Then
				' Cannot use long UNC paths
				Replace strPath, "\\?\", "\\"
			End If
			
			boolUNCPath = False
			
			'The number of items in the array that should be ignored because it is a drive mapping
			intServerDriveItems = 1
		End If
		
		On Error Resume Next
		
		If (Not m_objFSO.FileExists(strPath)) Then
			
			If (Err.Number <> 0) Then
				SetError Err.Number, Err.Description
				CreateParentFolder = Err.Number
				Exit Function
			End If
			
			arrPathParts = Split(strPath, "\")
			
			If (boolUNCPath) Then
				'Restore UNC path format
				strParentFolderPart = "\\"
			End If
			
			'Loop through array items of path parts excluding empty string items and file itself
			For intCountParts = 0 To (Ubound(arrPathParts) - 1)
				'Splitting path will create some empty string items that should be ignored
				If (arrPathParts(intCountParts) <> "") Then	
					
					strParentFolderPart = strParentFolderPart & arrPathParts(intCountParts) & "\"
					
					'Exclude \\server or drive letter from this check as it will always be False
					If (intCountParts > (intServerDriveItems)) Then
						If (Not m_objFSO.FolderExists(strParentFolderPart)) Then
							
							m_objFSO.CreateFolder(strParentFolderPart)
							
							If (Err.Number <> 0) Then
								SetError Err.Number, Err.Description
								CreateParentFolder = Err.Number
								Exit Function
							End If
							
						End If
					End If
					
				End If
			Next
			
		End If
		
		On Error Goto 0
		
		'Clear any previous errors as this method executed Successfully
		ClearError
		
		CreateParentFolder = 0
	End Function
	
	'Resets values for m_lngErrorLevel and m_strErrorDescription
	Private Function ClearError()
		SetError 0, ""
	End Function
	
	'Sets values for the m_lngErrorLevel and m_strErrorDescription properties
	Private Function SetError(lngErrorLevel, strErrorDescription)
		m_lngErrorLevel = lngErrorLevel
		m_strErrorDescription = strErrorDescription
	End Function
End Class
