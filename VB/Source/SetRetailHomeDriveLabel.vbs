'==========================================================================
'
' NAME: SetRetailHomeDriveLabel.vbs
' 
' VERSIONS:
'-------------------------------------------------------------
' 	VERSION	: 1.0
' 	AUTHOR	: Toru Ozawa 
'	DATE  	: 06/08/2018
' 	COMMENT	: Windows 10 project (http://intranet.nomuranow.com/JEE/browse/W10-426)
'-------------------------------------------------------------
'
'==========================================================================

Option Explicit

SetRetailHomeDriveLabel(Main)

' --dummy subroutine to have isolation for variable / constant--
Function SetRetailHomeDriveLabel(Main)

   Const HOMEDRIVE = "V:"
   Const ORGNAMESTART = "documents"

   Dim oApl, oShell, strLabel

   ' --dummy do loop to exit easily--
'   On error resume next
   do
      Main.LogFile.Log "Retail HomeDrive : Starting to change label on homedrive.", True

      Set oApl = WScript.CreateObject("Shell.Application")
      Set oShell = WScript.CreateObject("Wscript.Shell")

      '---Check if the use is Retail?----
      if False = IsRetailUser then
         Main.LogFile.Log "Retail HomeDrive : The user does not belong to Retail OU. Terminating script execution.", True
         exit do
      end if

      if oApl.NameSpace(HOMEDRIVE) is nothing then
         Main.LogFile.Log "Retail HomeDrive : V drive does not exist. Terminating script execution.", True
         exit do
      end if

      strLabel = oApl.NameSpace(HOMEDRIVE).Self.Name
      Main.LogFile.Log "Retail HomeDrive : Current label is " & strLabel, True

      '---Check if current label is default?----
      if 1 = Instr(strLabel, ORGNAMESTART) then
         strLabel = "'Fs' ‚Ì" & oShell.ExpandEnvironmentStrings("%USERNAME%")
         oApl.NameSpace(HOMEDRIVE).Self.Name = strLabel
         Main.LogFile.Log "Retail HomeDrive : New label is now " & strLabel, True
      else
         Main.LogFile.Log "Retail HomeDrive : Change is not required", True
         exit do
      end if

      exit do
   loop
   On error goto 0

   Set oShell = Nothing
   Set oApl = Nothing

   Main.LogFile.Log "Retail HomeDrive : Completed", True

End Function

