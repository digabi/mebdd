'    This file is part of MEBDD.
'
'    MEBDD is free software: you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    any later version.
'
'    MEBDD is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License
'    along with MEBDD.  If not, see <http://www.gnu.org/licenses/>.
'
'
'    This file contains all subs and functions related to Windows OS.
'    This script is included by mebdd.hta.

Option Explicit

Const HKEY_CURRENT_USER =  &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002
Const KEY_QUERY_VALUE = &H0001
Const KEY_SET_VALUE = &H0002

' This global variable contains sting of last reported error
Dim Win_LastError


' Checks if function [func_name] exists. Returns TRUE if it does
' http://stackoverflow.com/questions/921364/is-there-any-way-to-check-to-see-if-a-vbscript-function-is-defined
Function Win_FunctionExists (strFunctionName)
	Dim boolResult, f
	
	boolResult = False 

	On Error Resume Next

	Set f = GetRef(strFunctionName)

	If Err.number = 0 Then
		boolResult = True
	End If 	
	
	On Error GoTo 0

	Win_FunctionExists = boolResult
End Function

' Calls upstrem logging function and if it does not exists, creates a dialog
Sub Win_LogMessage (strMessage)
	If Win_FunctionExists("LogMessage") Then
		LogMessage strMessage
	Else
		MsgBox strMessage, 0+64, "Global LogMessage() missing"
	End If
End Sub

Function Win_GetEnvironmentString(strVariableName)
	Dim objShell
	Set objShell = CreateObject("WScript.Shell")
	Win_GetEnvironmentString = objShell.ExpandEnvironmentStrings("%" & strVariableName & "%")
End Function

Function Win_ReadRegKeyHKCU (strRegPath, strRegKey)
	Dim oReg, strValue, bHasAccessRight
	
	Win_LastError = ""
	
	Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
	oReg.CheckAccess HKEY_CURRENT_USER, strRegPath, KEY_QUERY_VALUE, bHasAccessRight
	
	If bHasAccessRight Then
		oReg.GetStringValue HKEY_CURRENT_USER,strRegPath,strRegKey,strValue
		
		If IsEmpty(strValue) or IsNull(strValue) Then
			Win_ReadRegKeyHKCU = ""
			Win_LastError = "REGISTRY_IS_EMPTY"
		Else
			Win_ReadRegKeyHKCU = strValue
		End If
	Else
		Win_ReadRegKeyHKCU = ""
		Win_LastError = "COULD_NOT_ACCESS_REGISTRY_FOR_QUERY"
	End If
End Function

Function Win_WriteRegKeyHKCU (strRegPath, strRegKey, strValue)
	Dim oReg, bHasAccessRight

	Win_LastError = ""
	
	If IsNull(strRegPath) or strRegPath = "" Then
		Win_WriteRegKeyHKCU = FALSE
		Win_LastError = "REGISTRY_PATH_IS_EMPTY"
		Exit Function
	End If
	
	If IsNull(strRegKey) or strRegKey = "" Then
		Win_WriteRegKeyHKCU = FALSE
		Win_LastError = "REGISTRY_KEY_IS_EMPTY"
		Exit Function
	End If
	
	Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
	
	' If the key exists check the access, otherwise just write
	If oReg.EnumKey(HKEY_CURRENT_USER, strRegPath, "", "") <> 0 Then
		' Key does not exist, create it
		oReg.CreateKey HKEY_CURRENT_USER, strRegPath
	End If
	
	oReg.CheckAccess HKEY_CURRENT_USER, strRegPath, KEY_SET_VALUE, bHasAccessRight

	If bHasAccessRight Then
		oReg.SetStringValue HKEY_CURRENT_USER, strRegPath, strRegKey, strValue
		Win_WriteRegKeyHKCU = True
	Else
		Win_WriteRegKeyHKCU = False
		Win_LastError = "REGISTRY_ACCESS_DENIED"
	End If
End Function

Function Win_ReadLastImageTag ()
	Win_ReadLastImageTag = Win_ReadRegKeyHKCU(REG_HKCU_PATH, REG_LAST_TAG_KEY)
End Function

Function Win_WriteLastImageTag (strNewTag)
	Win_WriteLastImageTag = Win_WriteRegKeyHKCU(REG_HKCU_PATH, REG_LAST_TAG_KEY, strNewTag)
End Function

' Checks if given file exists. Returns TRUE if file exists, otherwise FALSE
Function Win_FileExists (strFilePath)
	Dim objFSO
	
	Win_LastError = ""
	
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	If (objFSO.FileExists(strFilePath)) Then
		Win_FileExists = True
	Else
		Win_FileExists = False
	End If
End Function

' Checks if given folder exists. Return TRUE if folder exists, otherwise FALSE
Function Win_FolderExists (strPath)
	Dim objFSO
	
	Win_LastError = ""
	
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	If (objFSO.FolderExists(strFilePath)) Then
		Win_FolderExists = True
	Else
		Win_FolderExists = False
	End If
End Function

' Returns last extension from path or URL ("https://bubba.foo/something/file.bug" -> "bug")
' If no dot is found returns an empty string ("https://bubba.foo/something/filebug" -> "")
' We're not using the Windows GetExtensionName() as we need this to work with URLs
Function Win_GetFileExtension (strPath)
	Dim intDot, strExt
	
	intDot = InStrRev(strPath, ".")
	If IsNull(intDot) or IsEmpty(intDot) Then
		strExt = ""
	Elseif intDot = 0 Then
		strExt = ""
	Else
		strExt = Mid(strPath, intDot+1)
	End If
	
	Win_GetFileExtension = strExt
End Function

' Returns the basename of the path or URL
' "https://bubba.foo/something/file.bug" -> "file.bug"
' "\\server\path\filename.ext" -> "filename.ext"
' "C:/server/path/filename.ext" -> "filename.ext"
' If no slash is found returns the given string ("filebug.ext" -> "filebug.ext")
' We're not using the Windows GetBaseName() as we need this to work with URLs
Function Win_GetBaseName (strPath)
	Dim intSlash, strName
	
	intSlash = InStrRev(strPath, "/")
	If InStrRev(strPath, "\") > intSlash Then
		intSlash = InStrRev(strPath, "\")
	End If
	
	If IsNull(intSlash) or IsEmpty(intSlash) Then
		strName = strPath
	Elseif intSlash = 0 Then
		strName = strPath
	Else
		strName = Mid(strPath, intSlash+1)
	End If
	
	Win_GetBaseName = strName
End Function

' Returns the pathname of the path or URL
' "https://bubba.foo/something/file.bug" -> "https://bubba.foo/something"
' "\\server\path\filename.ext" -> "\\server\path\"
' "C:/server/path/filename.ext" -> "C:/server/path"
' If no slash is found returns the given string ("filebug.ext" -> "filebug.ext")
' We're not using the Windows GetAbsolutePathName() as we need this to work with URLs
Function Win_GetPathName (strPath)
	Dim intSlash, strName
	
	intSlash = InStrRev(strPath, "/")
	If InStrRev(strPath, "\") > intSlash Then
		intSlash = InStrRev(strPath, "\")
	End If
	
	If IsNull(intSlash) or IsEmpty(intSlash) Then
		strName = strPath
	Elseif intSlash = 0 Then
		strName = strPath
	Else
		strName = Left(strPath, intSlash-1)
	End If
	
	Win_GetPathName = strName
End Function

' Get temporary file name. This does not create the temporary file.
Function Win_GetTempFilename ()
	Dim objFSO
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Const intTemporaryFolder = 2
	
	Win_GetTempFilename = objFSO.GetSpecialFolder(intTemporaryFolder) & "\mebdd_" & objFSO.GetTempName
End Function


Function Win_GetLocalImageDirectoryPath ()
	Dim fso, strDestinationPath
	
	strDestinationPath = Win_GetEnvironmentString("LOCALAPPDATA") & "\" & LOCALDATA_PATH & "\"
	
	Set fso = CreateObject("Scripting.FileSystemObject") 
	If (Not fso.FolderExists(strDestinationPath)) Then
		Win_CreateDirs(strDestinationPath)
	End If
	
	Win_GetLocalImageDirectoryPath = strDestinationPath
End Function

Function Win_GetMEBDDInstallationPath ()
	Dim oReg, strValue, bHasAccessRight

	Win_LastError = ""
	
	Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")

	oReg.CheckAccess HKEY_LOCAL_MACHINE, REG_INSTALL_PATH, KEY_QUERY_VALUE, bHasAccessRight
	
	If bHasAccessRight = False Then
		MsgBox "En voi lukea MEBDD-asennushakemiston osoitetta. Asenna ohjelma uudelleen."
		Win_LastError = "COULD_NOT_ACCESS_REGISTRY_FOR_QUERY"
	End If
	
	oReg.GetStringValue HKEY_LOCAL_MACHINE,REG_INSTALL_PATH,REG_INSTALL_KEY,strValue
	
	If IsNull(strValue) Then
		Win_GetMEBDDInstallationPath = ""
	Else
		Win_GetMEBDDInstallationPath = strValue
	End If
End Function

' http://www.robvanderwoude.com/vbstech_folders_md.php
Sub Win_CreateDirs( MyDirName )
' This subroutine creates multiple folders like CMD.EXE's internal MD command.
' By default VBScript can only create one level of folders at a time (blows
' up otherwise!).
'
' Argument:
' MyDirName   [string]   folder(s) to be created, single or
'                        multi level, absolute or relative,
'                        "d:\folder\subfolder" format or UNC
'
' Written by Todd Reeves
' Modified by Rob van der Woude
' http://www.robvanderwoude.com

	Dim arrDirs, i, idxFirst, objFSO, strDir, strDirBuild

	' Create a file system object
	Set objFSO = CreateObject( "Scripting.FileSystemObject" )

	' Convert relative to absolute path
	strDir = objFSO.GetAbsolutePathName( MyDirName )

	' Split a multi level path in its "components"
	arrDirs = Split( strDir, "\" )

	' Check if the absolute path is UNC or not
	If Left( strDir, 2 ) = "\\" Then
		strDirBuild = "\\" & arrDirs(2) & "\" & arrDirs(3) & "\"
		idxFirst    = 4
	Else
		strDirBuild = arrDirs(0) & "\"
		idxFirst    = 1
	End If

	' Check each (sub)folder and create it if it doesn't exist
	For i = idxFirst to Ubound( arrDirs )
		strDirBuild = objFSO.BuildPath( strDirBuild, arrDirs(i) )
		If Not objFSO.FolderExists( strDirBuild ) Then
			objFSO.CreateFolder strDirBuild
		End if
	Next

	' Release the file system object
	Set objFSO= Nothing
End Sub

Function Win_DialogImageFile (strDialogHeader, strStartingPath, strFileSpec)
	Dim s, i2, strResultPath
	
	' Returns empty string if no file was selected
	strResultPath = ""
	
	s = Dlg.openfiledlg(CStr(strStartingPath & strFileSpec), , , CStr(strDialogHeader))  
	If (Len(s) > 0) Then
		' A file was selected
		
		'--strange HTMLDlgHelper behavior. Returns a string ending with nulls.
		'-- The nulls won't affect using the string, but they will matter if you test the string.
		'-- For instance: If UCase(Right(s, 3)) = "TXT" Then ....   That won't work unless the nulls are snipped.    
		'-- so check for nulls. If first null is first character that will return "". Otherwise there's a path string to
		'--  extract from the string buffer.
		i2 = InStr(s, Chr(0))
		If i2 > 1 Then  s = Left(s, (i2 - 1))
		' We're done, set the temporary result value
		strResultPath = s
	End If
	
	Win_DialogImageFile = strResultPath
End Function

' Delete given file [strFilePath]. Returns FALSE if the file still
' exists after delete.
Function Win_DeleteFile (strFilePath)
	Dim objFSO
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	If objFSO.FileExists(strFilePath) Then 
		objFSO.DeleteFile strFilePath
	End If
	
	If objFSO.FileExists(strFilePath) Then
		Win_DeleteFile = False
	Else
		Win_DeleteFile = True
	End If
End Function

' Delete given folder (strPath). Returns FALSE if the folder still
' exists after delete.
' The function does not distinguish between folders that have contents and those that do not.
' The specified folder is deleted regardless of whether or not it has contents.
Function Win_DeleteFolder (strPath)
	Dim objFSO
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	If objFSO.FolderExists(strPath) Then
		objFSO.DeleteFolder strPath
	End If
	
	If objFSO.FolderExists(strPath) Then
		Win_DeleteFolder = False
	Else
		Win_DeleteFolder = True
	End If
End Function

' Move given file (strSourcePath) to another location (strDestinationPath).
' Returns FALSE if given strSourcePath does not exists. If move fails creates
' a run-time error.
Function Win_MoveFile (strSourcePath, strDestinationPath)
	Dim objFSO
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	If Win_FileExists(strSourcePath) Then
		Win_MoveFile = True
		objFSO.MoveFile strSourcePath, strDestinationPath
	Else
		Win_MoveFile = False
	End If
End Function

' Recursively kills all subprocesses of the current process
' Returns True if all subprocesses were killed successfully
Function Win_KillMySubprocesses
	Dim PID
	PID = Win_GetCurrentProcessId
	Win_KillMySubprocesses = Win_KillProcesses(PID, False)
End Function

' Recursively kills all subprocesses of the given process (intPID)
' If boolKillParent is true kills the intPID as well
' Returns True if all subprocesses were killed successfully
' http://stackoverflow.com/questions/20379723/how-to-kill-child-processes-with-vbscript
Function Win_KillProcesses (intPID, boolKillParent)
	Dim PID, intResult, boolResult
	Dim objWMIService, colProcessList, objProcess
		
	PID = Win_GetCurrentProcessId
	Win_LastError = ""

	' Check for subprocesses kill them first
	Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\.\root\cimv2")
	Set colProcessList = objWMIService.ExecQuery("SELECT * FROM Win32_Process WHERE ParentProcessId = " & intPID)
	boolResult = True
	For Each objProcess in colProcessList 
		' Call myself with subprocess PID
		If not Win_KillProcesses(objProcess.ProcessId, True) Then
			' Send error downstream
			boolResult = False
		End If
	Next
	
	' Kill the given process
	If boolKillParent Then
		Set colProcessList = objWMIService.ExecQuery("SELECT * FROM Win32_Process WHERE ProcessId = " & intPID)
		For Each objProcess in colProcessList 
			intResult = objProcess.Terminate
			If intResult <> 0 Then
				Win_LogMessage "Failed to terminate process #" & objProcess.ProcessId & " (" & objProcess.Name & ")"
				Win_LastError = "KILL_FAILED"
				boolResult = False
			Else
				Win_LogMessage "Terminated process #" & objProcess.ProcessId & " (" & objProcess.Name & ")"
			End If
		Next
	End If
	
	Win_KillProcesses = boolResult
End Function

' Get PID of the current subprocess
' http://stackoverflow.com/questions/20379723/how-to-kill-child-processes-with-vbscript
Function Win_GetCurrentProcessId
	Dim oShell, sCmd, oWMI, oChldPrcs, oCols, lOut, lResult
	lOut = 0
	Set oShell  = CreateObject("WScript.Shell")
	Set oWMI    = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
	sCmd = "/K @echo " & Int(Rnd * 3333) * CDbl(Timer) \ 1
	oShell.Run "%comspec% " & sCmd, 0
	BEnd_Wait(100)
	Set oChldPrcs = oWMI.ExecQuery("Select * From Win32_Process Where CommandLine Like '%" & sCmd & "'", ,32)
	For Each oCols In oChldPrcs
		lOut = oCols.ParentProcessId
		lResult = oCols.Terminate
		Exit For
	Next
	Set oChldPrcs = Nothing
	Set oWMI = Nothing
	Set oShell = Nothing
	Win_GetCurrentProcessId = lOut
End Function
