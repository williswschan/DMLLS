Option Explicit

Dim Main
Set Main = New DesktopManagement

Class DesktopManagement
	'Logging Class
	Private m_objLogFile
	'ComputerClass
	Private m_objComputer
	'User Class
	Private m_objUser
	'Script Version
	Private m_strScriptVer
	
	'String to identify the type of job that is being run
	Private m_strJobType
	'Location of the log file parent folder
	Private m_strLogFileFolder
	'Registry path for further logging
	Private m_strRegistryPath
	
	'Indicates if verbose logging has been specified
	Private m_boolVerboseLogging
	'Max age (in days) of log files to retain 
	Private m_intMaxLogAge
	'WMI Date time object
	Private m_dateTime_wmi
	'Script start and end time (in WMI and regular DateTime Format)
	Private m_dateInitialTime_wmi
	Private m_dateEndTime_wmi
	Private m_dateInitialTime_RegularTime
	Private m_dateEndTime_RegularTime
			
	'LogFile object Property
	Public Property Get LogFile
		Set LogFile = m_objLogFile
	End Property
	
	Public Property Let LogFile(objLogFile)
		Set m_objLogFile = objLogFile
	End Property
	
	'Script Version Property (read only)
	Public Property Get ScriptVersion
		Set ScriptVersion = m_strScriptVer
	End Property
	
	'Computer object Property (read only)
	Public Property Get Computer
		Set Computer = m_objComputer
	End Property
	
	'User object property (read only)
	Public Property Get User
		Set User = m_objUser
	End Property
	
	'JobType string property (read only)
	Public Property Get JobType
		JobType = m_strJobType
	End Property
	
	'LogFilePath string property (read only)
	Public Property Get LogFilePath
		LogFilePath = m_strLogFileFolder
	End Property
	
	'RegistryPath string property (read only)
	Public Property Get RegistryPath
		RegistryPath = m_strRegistryPath
	End Property
	
	
	'Constructor. Set the default values
	Private Sub Class_Initialize()
		Dim objWshShell
		
		Set objWshShell = CreateObject("WScript.Shell")
		Set m_dateTime_wmi = CreateObject("WbemScripting.SWbemDateTime")		
		
		'Set script start time
		m_dateTime_wmi.SetVarDate(Now)
		m_dateInitialTime_wmi = m_dateTime_wmi
		m_dateInitialTime_RegularTime = m_dateTime_wmi.GetVarDate
		
		'Create instance of dependency classes. These must be loaded into script via wsf
		Set m_objLogFile	= New LoggingObject		'Logging.vbs
		Set m_objComputer	= New ComputerObject	'Machine.vbs
		Set m_objUser 		= New UserObject		'User.vbs
		
		'Load resources from wsf
		m_strJobType 		= getResource("JobType")
		m_strLogFileFolder 	= objWshShell.ExpandEnvironmentStrings(getResource("LogFilePath"))
		m_strRegistryPath 	= getResource("RegistryPath")
		
		'Version Control - must be in String type.
		m_strScriptVer = "1.29"
				
		'Get values for arguments
		If (GetNamedArgumentValue("Verbose" ,"False") <> "False") Then
			m_objLogFile.VerboseLogging = True
		Else
			m_objLogFile.VerboseLogging = False
		End If
		
		m_intMaxLogAge = CInt(GetNamedArgumentValue("MaxLogAge" ,"60"))
		
		'Set path for log file
		m_objLogFile.Path = m_strLogFileFolder & "\" & m_strJobType & "_" & m_objComputer.Name & "_" & m_objLogFile.TimeStamp(m_dateInitialTime_RegularTime, False) & ".Log"
		
		'Begin creating Log contents
		m_objLogFile.Update "Log for " & WScript.ScriptName & " initialised on " & m_dateInitialTime_RegularTime & vbCrLf & "---" & vbCrLf & _
							"Job Type:       " & vbTab & m_strJobType & vbCrLf & _
							"Log File Path:  " & vbTab & m_objLogFile.Path & vbCrLf & _
							"Registry Path:  " & vbTab & m_strRegistryPath & vbCrLf & _
							"Verbose Logging:" & vbTab & m_objLogFile.VerboseLogging & vbCrLf & _
							"Log Retention:  " & vbTab & m_intMaxLogAge & " days" & vbCrLf & _
							"Script Path:    " & vbTab & WScript.ScriptFullName & VbCrLf & _
							"Script Version:    " & vbTab & m_strScriptVer & VbCrLf & VbCrLf
														
		If (m_objLogFile.VerboseLogging) Then
			m_objLogFile.Update "Computer" & vbCrLf & "---" & vbCrLf & _
							m_objComputer.ToString() & vbCrLf & _
							"User" & vbCrLf & "---" & vbCrLf & _ 
							m_objUser.ToString()
		Else
			m_objLogFile.Update "Computer Name:   " & vbTab & m_objComputer.Name & vbCrLf & _
							"Computer Domain: " & vbTab & m_objComputer.Domain & vbCrLf & _
							"User Name:       " & vbTab & m_objUser.Name & vbCrLf & _
							"User Domain:     " & vbTab & m_objUser.Domain & vbCrLf
		End If
		
		m_objLogFile.Log "Initialize: Completed initialization of " & WScript.ScriptName & " " & m_strJobType & " job", True
	End Sub
	
	'Destructor
	Private Sub Class_Terminate()
		Dim objWshShell, objFSO, objFolder, objFile
		Dim strRegistryJobTypePath, strFilePath
		Dim intPurgedLogs, intRunTime

		
		m_objLogFile.Log "Finalize: Finalizing the " & m_strJobType & " job", True
		
		'Registry path defined in the wsf file will be appended with 'JobType -' to create unique value names for each JobType
		strRegistryJobTypePath = m_strRegistryPath & "\" & m_strJobType & " - "
		
		Set objWshShell = CreateObject("WScript.Shell")
		Set objFSO = CreateObject("Scripting.FileSystemObject")
		
		intPurgedLogs = 0
		
		If (objFSO.FolderExists(m_strLogFileFolder)) Then
			Set objFolder = objFSO.GetFolder(m_strLogFileFolder)
			
			'Delete any files that meet name and date criteria
			For Each objFile in objFolder.Files
				If ((Left(Ucase(objFile.Name), Len(m_strJobType)) = UCase(m_strJobType)) And _
					(UCase(Right(objFile.Name, 4)) = ".LOG") And _
					(DateDiff("d", objFile.DateCreated, Now) > m_intMaxLogAge)) Then
					
					On Error Resume Next
					
					'Store the path for logging before object is deleted
					strFilePath = objFile.Path
					
					objFile.Delete True
					
					If (Err.Number = 0) Then
						m_objLogFile.Log "Finalize: Purged log file '" & strFilePath & "'", False
						intPurgedLogs = intPurgedLogs + 1
					Else
						m_objLogFile.Log "Finalize: Failed to purge log file '" & strFilePath & "'. " & Err.Description & "(" & Err.Number & ")", True
					End If
					
					On Error Goto 0
				End If
			Next
			
			m_objLogFile.Log "Finalize: Total of " & intPurgedLogs & " log files purged", True
		End If
		
		'Set script end time
		m_dateTime_wmi.SetVarDate(Now)
		m_dateEndTime_wmi = m_dateTime_wmi
		m_dateEndTime_RegularTime = m_dateTime_wmi.GetVarDate
 		intRunTime = DateDiff("s", m_dateInitialTime_RegularTime, m_dateEndTime_RegularTime)
		
		m_objLogFile.Log "Finalize: Registry logging path '" & strRegistryJobTypePath & "'", False
		
		On Error Resume Next
		
		'Record basic information in user registry to help quickly query environment (as opposed ot parsing all log files)
		objWshShell.RegWrite strRegistryJobTypePath & "Script Name", WScript.ScriptFullName, "REG_SZ"
		objWshShell.RegWrite strRegistryJobTypePath & "Log File", m_objLogFile.Path, "REG_SZ"
		objWshShell.RegWrite strRegistryJobTypePath & "Start Time", m_dateInitialTime_wmi, "REG_SZ"
		objWshShell.RegWrite strRegistryJobTypePath & "End Time", m_dateEndTime_wmi, "REG_SZ"
		objWshShell.RegWrite strRegistryJobTypePath & "Run Time (seconds)", intRunTime, "REG_SZ"
		objWshShell.RegWrite strRegistryJobTypePath & "Script Version", m_strScriptVer, "REG_SZ"
		
		If (Err.Number = 0) Then
			m_objLogFile.Log "Finalize: Registry updated successfully", False
		Else
			m_objLogFile.Log "Finalize: Error updating registry. " & Err.Description & "(" & Err.Number & ")", True
		End If
		
		On Error Goto 0
		
		'Finalise log file contents
		m_objLogFile.Update vbCrLf & "---" & vbCrLf & WScript.ScriptName & " completed. Total run time: " & intRunTime & " seconds"
		'Append contents to log file
		m_objLogFile.AppendContentToFile
	End Sub
	
	'Get named argument value else return a default value
	Function GetNamedArgumentValue(strArgName, strDefaultValue)
		If (WScript.Arguments.Named.Exists(strArgName)) Then
			GetNamedArgumentValue = WScript.Arguments.Named.Item(strArgName)
		Else
			GetNamedArgumentValue = CStr(strDefaultValue)
		End If
	End Function
End Class
