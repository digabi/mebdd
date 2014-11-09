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
'    This file contains all subs and functions related to MEBDD functions
'    like downloading files, counting MD5 digests and calling mebdd_worker
'    to write images. This script is included by mebdd.hta.

' This global variable contains sting of last reported error
Dim BEnd_LastError


' Checks if function [func_name] exists. Returns TRUE if it does
' http://stackoverflow.com/questions/921364/is-there-any-way-to-check-to-see-if-a-vbscript-function-is-defined
Function BEnd_FunctionExists (strFunctionName)
	Dim boolResult, f
	
	boolResult = False 

	On Error Resume Next

	Set f = GetRef(strFunctionName)

	If Err.number = 0 Then
		boolResult = True
	End If 	
	
	On Error GoTo 0

	BEnd_FunctionExists = boolResult
End Function

' Calls upstrem logging function and if it does not exists, creates a dialog
Sub BEnd_LogMessage (strMessage)
	If BEnd_FunctionExists("LogMessage") Then
		LogMessage strMessage
	Else
		MsgBox strMessage, 0+64, "Global LogMessage() missing"
	End If
End Sub

Function BEnd_FileMD5IsEqual(strFilePath, strMD5)
	Dim strMD5Cmd, exitcode, boolResult
	
	boolResult = False
	BEnd_LastError = ""

	strMD5Cmd = Chr(34) & BIN_MEBMD5 & Chr(34) & " " & Chr(34) & strFilePath & Chr(34) & " " & strMD5
	
	BEnd_LogMessage "MD5SUM command: " & strMD5Cmd
	
	Dim shell
	Set shell = CreateObject("WScript.Shell")
	exitcode = shell.Run(strMD5Cmd, 0, TRUE)

	If exitcode = 0 Then
		' Error in MEBMD5
		BEnd_LogMessage "Varoitus! MEBMD5-komento tuotti virheen: " & exitcode
	End If
	
	If exitcode = 1 Then
		' Sums match
		boolResult = True
	End If
	
	BEnd_FileMD5IsEqual = boolResult
End Function

Function BEnd_DownloadImage (strURL, strDestinationFile)
	Dim strCurlCmd, exitcode, boolResult
	
	boolResult = False
	BEnd_LastError = ""
	
	strCurlCmd = Chr(34) & BIN_CURL & Chr(34) & " -L -k """ & strURL & """ -o """ & strDestinationFile & Chr(34)
	
	Dim shell
	Set shell = CreateObject("WScript.Shell")
	BEnd_LogMessage "calling cURL: " & strCurlCmd
	
	exitcode = shell.Run(strCurlCmd, 7, TRUE)
	
	BEnd_LogMessage "cURL returns: " & exitcode 
	
	If (exitcode = 0) Then
		boolResult = True
	Else
		BEnd_LastError = exitcode
	End If
	
	BEnd_DownloadImage = boolResult
End Function

Function BEnd_DownloadMD5 (strURL)
	Dim objBrowser, strMD5, intLoops
	Const intTimeoutLoops = 10
	
	'Set objBrowser = CreateObject("MSXML2.XMLHTTP")
	Set objBrowser = CreateObject("MSXML2.ServerXMLHTTP.6.0")
	objBrowser.open "GET", strURL, True
	' FIXME: This dies if server is not known
	If not IsNull(objBrowser) and not IsEmpty(objBrowser) Then
		' We have the browser object
		objBrowser.send
		
		intLoops = 0
		While objBrowser.readyState <> 4
			On Error Resume Next
			objBrowser.waitForResponse(1000)
			If Err.Number <> 0 Then
				' Browser returns error code
				strMD5 = ""
				BEnd_LastError = "BROWSER_ERROR"
				Exit Function
			End If
			On Error Goto 0

			' Check timeout
			intLoops = intLoops+1
			If intLoops > intTimeoutLoops Then
				' Timeout
				strMD5 = ""
				BEnd_LastError = "TIMEOUT"
				Exit Function
			End If
		Wend
	
		If objBrowser.Status = 200 Then
			' Everything was OK
			strMD5 = objBrowser.responseText
		Else
			strMD5 = ""
			BEnd_LastError = objBrowser.Status
		End If
	Else
		' No browser object
		strMD5 = ""
		BEnd_LastError = "COULD_NOT_CONNECT_SERVER"
	End If
	
	' We expect the MD5 to be the first part of the string
	strMD5 = Trim(strMD5)
	If (InStr(strMD5, " ") > 0) Then
		' There is at least one space in the string, take the leftmost part
		strMD5 = Left(strMD5, InStr(strMD5, " ")-1)
	End If
	
	BEnd_DownloadMD5 = strMD5 
End Function

Function BEnd_DownloadMD5AndStore (strURL, strKeyPath, strKeyName)
	Dim boolResult, oBrowser, strMD5
	
	boolResult = False
	BEnd_LastError = ""
	
	strMD5 = BEnd_DownloadMD5(strURL)
	
	If strMD5 <> "" Then
		' We have a string (and expect it to be a MD5 digest)
	
		' Store MD5 to given registry
		If Win_WriteRegKeyHKCU(strKeyPath, strKeyName, strMD5) Then
			boolResult = True
		Else
			boolResult = False
			BEnd_LastError = "FAILED_TO_WRITE_REGISTRY"
		End If
	Else
		boolResult = False
		' We don't set the BEnd_LastError here as it contains the error set by BEnd_DownloadMD5
	End If
	
	BEnd_DownloadMD5AndStore = boolResult
End Function

Function BEnd_WriteImage (strImageFile, arrSelectedDrives, boolVerifyImage)
	Dim strWorkerCmd, exitcode
	
	BEnd_LastError = ""

	' Check if file exists
	Dim fso
	Set fso = CreateObject("Scripting.FileSystemObject")
	If (Not fso.FileExists(strImageFile)) Then
		BEnd_LastError = "IMAGE_FILE_NOT_FOUND"
		BEnd_WriteImage = False
		Exit Function
	End If
	
	arrSelectedDrives = GetSelectedDrives()
	
	If (UBound(arrSelectedDrives) = 0) Then
		BEnd_LastError = "NO_DRIVES_SELECTED"
		BEnd_WriteImage = False
		Exit Function
	End If
	
	If boolVerifyImage Then
		' Execute mebdd_worker with verify
		strWorkerCmd = Chr(34) & BIN_MEBDD_WORKER & Chr(34) & " -v " & Chr(34) & strImageFile & Chr(34)
	Else
		' No verify this time
		strWorkerCmd = Chr(34) & BIN_MEBDD_WORKER & Chr(34) & " " & Chr(34) & strImageFile & Chr(34)
	End If
	
	For n = 1 to UBound(arrSelectedDrives)
		strWorkerCmd = strWorkerCmd & " " & Chr(34) & arrSelectedDrives(n) & Chr(34)
	Next
	
	Dim shell
	Set shell = CreateObject("WScript.Shell")
	
	BEnd_LogMessage "calling mebdd_worker: " & strWorkerCmd

	exitcode = shell.Run(strWorkerCmd, 7, TRUE)

	BEnd_LogMessage "mebdd_worker returns: " & exitcode
	
	If (exitcode = 255) Then
		BEnd_LastError = ""
		BEnd_WriteImage = True
	Else
		BEnd_LastError = exitcode
		BEnd_WriteImage = False
	End If
End Function

' Unbox archive with a given archive-specific low level function
' Returns true on success
Function BEnd_Unbox (strUnboxer, strPathBox, strPathLocal)
	If UCase(strUnBoxer) = "ZIP" Then
		BEnd_Unbox = BEnd_UnboxZip(strPathBox, strPathLocal)
	Else
		BEnd_LogMessage "Unknown unboxer: " & strUnboxer
		BEnd_LastError = "UNKNOWN_UNBOXER"
		BEnd_Unbox = False
	End If
End Function

' Unbox using Info-ZIP
' Returns true on success
Function BEnd_UnboxZip (strPathBox, strPathLocal)
	Dim strLocalPath, strLocalBasename, strUnzipCmd, exitcode
	
	strLocalBasename = Win_GetBaseName(strPathLocal)
	strLocalPathname = Win_GetPathName(strPathLocal)
	
	' Do we have a basename and a pathname?
	If strLocalBasename = "" Then
		' Nope, quit here
		BEnd_LastError = "ZIP_BASENAME_MISSING"
		BEnd_UnboxZip = False
		Exit Function
	End If
	If strLocalPathname = "" Then
		' Nope, quit here
		BEnd_LastError = "ZIP_PATHNAME_MISSING"
		BEnd_UnboxZip = False
		Exit Function
	End If
	
	strUnzipCmd = Chr(34) & BIN_UNZIP & Chr(34) & " -u " & Chr(34) & strPathBox & Chr(34) & " " & Chr(34) & strLocalBasename & Chr(34) & " -d " & Chr(34) & strLocalPathname & Chr(34)

	Dim objShell
	Set objShell = CreateObject("WScript.Shell")
	
	BEnd_LogMessage "calling unzip: " & strUnzipCmd

	exitcode = objShell.Run(strUnzipCmd, 0, TRUE)

	BEnd_LogMessage "unzip returns: " & exitcode
	
	If exitcode = 0 Then
		' Normal termination
		BEnd_UnboxZip = True
	Else
		BEnd_LastError = "ZIP_ERROR"
		BEnd_UnboxZip = False
	End If
End Function


' http://sogeeky.blogspot.fi/2007/04/vbscript-function-code-to-convert-bytes.html
Function BEnd_ConvertSize(Size)
	Dim CommaLocate, Suffix
	
	Do While InStr(Size,",") 'Remove commas from size
		CommaLocate = InStr(Size,",")
		Size = Mid(Size,1,CommaLocate - 1) & _
			Mid(Size,CommaLocate + 1,Len(Size) - CommaLocate)
	Loop

	Suffix = " Bytes"
	If Size >= 1024 Then suffix = " KB"
	If Size >= 1048576 Then suffix = " MB"
	If Size >= 1073741824 Then suffix = " GB"
	If Size >= 1099511627776 Then suffix = " TB"

	Select Case Suffix
		Case " KB" Size = Round(Size / 1024, 1)
		Case " MB" Size = Round(Size / 1048576, 1)
		Case " GB" Size = Round(Size / 1073741824, 1)
		Case " TB" Size = Round(Size / 1099511627776, 1)
	End Select

	BEnd_ConvertSize = Size & Suffix
End Function

' Returns TRUE if [item] is a member of array [A]
Function BEnd_InArray(item,A) 
     Dim i 
     For i=0 To UBound(A) Step 1 
         If A(i) = item Then 
             BEnd_InArray=True 
             Exit Function 
         End If 
     Next 
     BEnd_InArray=False 
 End Function 
