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

Option Explicit

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

' Enumerate all image sources
' Returns a hash array ("Scripting.Dictionary" object)
Function BEnd_GetImageSources (strIniFile)
	Dim objReturn, objIni, strSection, strKey
	Set objReturn = CreateObject("Scripting.Dictionary")
	Set objIni = BEnd_ReadIniFile(strIniFile)
	objReturn.RemoveAll()
	
	If objIni Is Nothing Then
		' Failed to load ini file
		' Return empty
		BEnd_LogMessage "Failed to load image .ini file: " & strIniFile
	Else
		' We have content
		' Loop through all sections
		For Each strSection in objIni.Keys()
			If Left(strSection, 1) = "*" Then
				' Section names beginning with * have special meaning
				If LCase(strSection) = "*system" Then
					' This section redefines some global system settings
					If objIni(strSection).Exists("imageupdatemethod") Then
						IMAGE_UPDATE_METHOD = objIni(strSection)("imageupdatemethod")
					End If
					If objIni(strSection).Exists("imageversionurl") Then
						IMAGE_VERSION_URL = objIni(strSection)("imageversionurl")
					End If
				End If
			Else
				' This section is a normal one and defines an image
				
				' REG_HKCU_PATH is a global variable
				objReturn.Add strSection, Array(_
					objIni(strSection)("legend"),_
					objIni(strSection)("description"),_
					objIni(strSection)("urlimage"),_
					objIni(strSection)("urlmd5"),_
					Win_GetLocalImageDirectoryPath() & objIni(strSection)("localfile"),_
					REG_HKCU_PATH,_
					objIni(strSection)("regremotemd5"),_
					objIni(strSection)("reglocalmd5"),_
					objIni(strSection)("reglocalversion")_
					)
			End If
		Next
	End If
	
	Set BEnd_GetImageSources = objReturn
End Function

' Reads Windows-style .ini file and returns a hash array ("Scripting.Dictionary" object)
' Returns Nothing if given file does not exist
' http://stackoverflow.com/questions/21825192/read-data-from-ini-file
Function BEnd_ReadIniFile(sFSpec)
	Dim dicTmp, tsIn, objFSO
	
	If not Win_FileExists(sFSpec) Then
		' Given ini not found
		Set BEnd_ReadIniFile = Nothing
		Exit Function
	End If
	
	Set dicTmp = CreateObject("Scripting.Dictionary")
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set tsIn = objFSO.OpenTextFile(sFSpec)
	
	Dim sLine, sSec, aKV
	Do Until tsIn.AtEndOfStream
		sLine = tsIn.ReadLine()
		' Replace tab with space
		sLine = Replace(sLine, Chr(9), " ")
		sLine = Trim(sLine)
		If ";" = Left(sLine, 1) Or "#" = Left(sLine, 1) Then
			' This is remark, do nothing
		ElseIf "[" = Left(sLine, 1) Then
			sSec = Mid(sLine, 2, Len(sLine) - 2)
			Set dicTmp(sSEc) = CreateObject("Scripting.Dictionary")
		Else
			If "" <> sLine Then
				aKV = Split(sLine, "=")
				If 1 = UBound(aKV) Then
					dicTmp(sSec)(Trim(LCase(aKV(0)))) = Trim(aKV(1))
				End If
			End If
		End If
	Loop
	
	tsIn.Close
	Set BEnd_ReadIniFile = dicTmp
End Function

' Check if file strFilePath as MD5 sum strMD5.
' Returns:
' - True: The MD5 of the file equals to given strMD5
' - False: The file has different MD5
Function BEnd_FileMD5IsEqual(strFilePath, strMD5)
	Dim strMD5Cmd, exitcode, boolResult
	
	boolResult = False
	BEnd_LastError = ""
	
	If not Win_FileExists(strFilePath) Then
		BEnd_FileMD5IsEqual = False
		BEnd_LastError = "MD5SUM_FILE_MISSING"
		Exit Function
	End If

	strMD5Cmd = Chr(34) & BIN_MEBMD5 & Chr(34) & " " & Chr(34) & strFilePath & Chr(34) & " " & strMD5
	
	BEnd_LogMessage "MD5SUM command: " & strMD5Cmd
	
	Dim shell
	Set shell = CreateObject("WScript.Shell")
	exitcode = shell.Run(strMD5Cmd, 0, TRUE)

	If exitcode = 0 Then
		' Error in MEBMD5
		BEnd_LogMessage "Warning! MEBMD5 exits with exitcode ZERO, maybe an unhandled exception?"
		boolResult = False
		BEnd_LastError = "MD5SUM_EXCEPTION"
	End If
	
	If exitcode = 1 Then
		' Sums match
		boolResult = True
	End If
	
	BEnd_FileMD5IsEqual = boolResult
End Function

Function BEnd_DownloadImage (strURL, strDestinationFile)
	Dim strCurlCmd, exitcode, boolResult, boolRetry
	
	boolResult = False
	BEnd_LastError = ""
	
	strCurlCmd = Chr(34) & BIN_CURL & Chr(34) & " -L -k --speed-time 5 --speed-limit 256 -C - """ & strURL & """ -o """ & strDestinationFile & Chr(34)
	
	Dim shell
	Set shell = CreateObject("WScript.Shell")
	
	boolRetry = True
	While boolRetry
		BEnd_LogMessage "calling cURL: " & strCurlCmd
		
		exitcode = shell.Run(strCurlCmd, 7, TRUE)
		
		BEnd_LogMessage "cURL returns: " & exitcode 
		If exitcode = 56 Or exitcode = 28 Then
			' Timeout, retry...
			boolRetry = True
			BEnd_LogMessage "Got timeout (#" & exitcode & "), retrying"
		Else
			' Other error, fail...
			boolRetry = False
		End If
	Wend
	
	If (exitcode = 0) Then
		boolResult = True
	Else
		BEnd_LastError = "CURL_ERROR_" & exitcode
	End If
	
	BEnd_DownloadImage = boolResult
End Function

Function BEnd_DownloadMD5 (strURL)
	Dim strMD5
	strMD5 = BEnd_DownloadText(strURL)
	
	' We expect the MD5 to be the first part of the string
	strMD5 = Trim(strMD5)
	If (InStr(strMD5, " ") > 0) Then
		' There is at least one space in the string, take the leftmost part
		strMD5 = Left(strMD5, InStr(strMD5, " ")-1)
	End If
	
	BEnd_DownloadMD5 = strMD5 
End Function

Function BEnd_DownloadVersion (strURL)
	Dim strVersion, objRE
	Set objRE = New RegExp
	
	strVersion = BEnd_DownloadText(strURL)
	
	' Accept only digits and commas
	objRE.Global = True
	objRE.Pattern = "[^\d]"
	strVersion = objRE.Replace(strVersion, "")
	
	BEnd_DownloadVersion = strVersion
End Function

Function BEnd_DownloadText (strURL)
	Dim objBrowser, strResult, intLoops
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
				BEnd_DownloadText = ""
				BEnd_LastError = "BROWSER_ERROR"
				Exit Function
			End If
			On Error Goto 0

			' Check timeout
			intLoops = intLoops+1
			If intLoops > intTimeoutLoops Then
				' Timeout
				BEnd_DownloadText = ""
				BEnd_LastError = "TIMEOUT"
				Exit Function
			End If
		Wend
	
		If objBrowser.Status = 200 Then
			' Everything was OK
			BEnd_DownloadText = objBrowser.responseText
		Else
			BEnd_DownloadText = ""
			BEnd_LastError = objBrowser.Status
		End If
	Else
		' No browser object
		BEnd_DownloadText = ""
		BEnd_LastError = "COULD_NOT_CONNECT_SERVER"
	End If
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

' Download image file from strUrlImage and image hash from strUrlMD5. Report
' progress to strCallbackFuncName and use strImageName (an user-friendly name of the
' image) when reporting.
' Returns and Scripting.Dictionary object with two values, "filename" and "md5"
' - on success:
'      "filename" contains a pathname to the verified image file
'      "md5" contains the hash value of the image
'               
' - on failure:
'      both values are set as an empty string + sets BEnd_LastError
Function BEnd_DownloadAndVerify(strCallbackFuncName, strUrlImage, strUrlMD5, strFinalImageFilename, strImageName)
	Dim objReportFunc, objResult, strDownloadedMD5, strLocalPath, strLocalFile, boolZIP, strZipPath, strZipFile, boolDownloadOK
	
	Set objResult = CreateObject("Scripting.Dictionary")
	objResult("filename") = ""
	objResult("md5") = ""
	
	' Get callback function object
	Set objReportFunc = GetRef(strCallbackFuncName)
	
	objReportFunc(BEnd_FMT(VBGT_Get("Downloading image checksum for image '%x'..."), Array(strImageName)))
	BEnd_LogMessage "Downloading image checksum from " & strUrlMD5
	
	' Download MD5
	strDownloadedMD5 = BEnd_DownloadMD5(strUrlMD5)
	
	If strDownloadedMD5 = "" Then
		' MD5 download failed
		BEnd_LogMessage "Failed to download MD5: " & BEnd_LastError
		' We do not set BEnd_LastMessage as we want to report that to the upstream
		Set BEnd_DownloadAndVerify = objResult
		Exit Function
	End If
	
	' Create a local temporary path to store files
	strLocalPath = Win_GetTempFilename()
	Win_CreateDirs(strLocalPath)
	
	' Are we downloading a zip?
	boolZIP = False
	If UCase(Win_GetFileExtension(strUrlImage)) = "ZIP" Then
		boolZIP = True
	End If
	
	' The Image will be downloaded to strLocalFile
	strLocalFile = strLocalPath & "\" & Win_GetBasename(Win_GetTempFilename())
	
	' Download image file
	objReportFunc(BEnd_FMT(VBGT_Get("Downloading image '%x' data from the server. This takes a long time..."), Array(strImageName)))
	boolDownloadOK = BEnd_DownloadImage(strUrlImage, strLocalFile)
	
	' If download failed, remove temporary stuff and exit
	If not boolDownloadOK Then
		If not Win_DeleteFolder(strLocalPath) Then
			BEnd_LogMessage "Warning: Failed to remove temporary folder " & strLocalPath
		End If
		' We do not set BEnd_LastMessage as we want to report that to the upstream
		Set BEnd_DownloadAndVerify = objResult
		Exit Function
	End If
	
	' Have we downloaded a zip file? Unpack it to another temp file
	If boolZIP Then
		' Store ZIP path
		strZipFile = strLocalFile
		
		' Create new temporary folder & filename
		strLocalPath = Win_GetTempFilename()
		Win_CreateDirs(strLocalPath)
		strLocalFile = strLocalPath & "\" & Win_GetBaseName(strFinalImageFilename)
		
		objReportFunc(BEnd_FMT(VBGT_Get("Unpacking the image '%x' data..."), Array(strImageName)))

		If not BEnd_Unbox("zip", strZipFile, strLocalFile) Then
			BEnd_LogMessage "Image unboxing failed, error: " & BEnd_LastError
			
			' Delete zip file
			If not Win_DeleteFile(strZipFile) Then
				LogMessage("Warning: Could not remove zip file " & strZipFile)
			End If

			BEnd_LastError = "UNBOXING_FAILED"
			Set BEnd_DownloadAndVerify = objResult
			Exit Function
		End If
		
		BEnd_LogMessage "The ZIP package has been unpacked from " & strZipFile & " to " & strLocalFile
		If not Win_DeleteFile(strZipFile) Then
			BEnd_LogMessage "Warning: Could not remove temporary zip file " & strZipFile
		End If
		If not Win_DeleteFolder(Win_GetPathName(strZipFile)) Then
			BEnd_LogMessage "Warning: Could not remove temporary zip directory " & Win_GetPathName(strZipFile)
		End If
	End If
	
	' Check MD5 of the downloaded and possibly unpacked file
	objReportFunc(BEnd_FMT(VBGT_Get("Checking image '%x' data..."), Array(strImageName)))
	If BEnd_FileMD5IsEqual(strLocalFile, strDownloadedMD5) Then
		' The MD5 matches
		objResult("filename") = strLocalFile
		objResult("md5") = strDownloadedMD5
		BEnd_LogMessage "MD5 check passed for file " & strLocalFile
	Else
		BEnd_LastError = "CHECKSUM_MISMATCH"
		BEnd_LogMessage "MD5 check failed for file " & strLocalFile
	End If
	
	Set BEnd_DownloadAndVerify = objResult
End Function

Function BEnd_WriteImage (strImageFile, arrSelectedDrives, boolVerifyImage)
	Dim strWorkerCmd, exitcode, errorcode, n, objShell
	
	BEnd_LastError = ""

	' Check if file exists
	Dim fso
	Set fso = CreateObject("Scripting.FileSystemObject")
	If (Not fso.FileExists(strImageFile)) Then
		BEnd_LastError = "IMAGE_FILE_NOT_FOUND"
		BEnd_WriteImage = False
		Exit Function
	End If
	
	If boolVerifyImage Then
		' Execute mebdd_worker with verify
		strWorkerCmd = Chr(34) & BIN_MEBDD_WORKER & Chr(34) & " -w -v " & Chr(34) & strImageFile & Chr(34)
	Else
		' No verify this time
		strWorkerCmd = Chr(34) & BIN_MEBDD_WORKER & Chr(34) & " -w " & Chr(34) & strImageFile & Chr(34)
	End If
	
	For n = 1 to UBound(arrSelectedDrives)
		strWorkerCmd = strWorkerCmd & " " & Chr(34) & arrSelectedDrives(n) & Chr(34)
	Next
	
	Set objShell = CreateObject("WScript.Shell")
	
	BEnd_LogMessage "calling mebdd_worker: " & strWorkerCmd

	On Error Resume Next
	exitcode = objShell.Run(strWorkerCmd, 7, TRUE)
	errorcode = Err.Number
	On Error Goto 0
	
	If errorcode <> 0 Then
		BEnd_LogMessage "mebdd_worker shell object returns error code: " & errorcode
		BEnd_LastError = ""
		BEnd_WriteImage = False
		Exit Function
	End If

	BEnd_LogMessage "mebdd_worker returns: " & exitcode
	
	If (exitcode = 255) Then
		BEnd_LastError = ""
		BEnd_WriteImage = True
	Else
		BEnd_LastError = exitcode
		BEnd_WriteImage = False
	End If
End Function

Function BEnd_CreateFilesystem (arrSelectedDrives)
	Dim strWorkerCmd, exitcode, errorcode, n, objShell
	
	BEnd_LastError = ""

	strWorkerCmd = Chr(34) & BIN_MEBDD_WORKER & Chr(34) & " -c"
	
	For n = 1 to UBound(arrSelectedDrives)
		strWorkerCmd = strWorkerCmd & " " & Chr(34) & arrSelectedDrives(n) & Chr(34)
	Next
	
	Set objShell = CreateObject("WScript.Shell")
	
	BEnd_LogMessage "calling mebdd_worker: " & strWorkerCmd

	On Error Resume Next 
	exitcode = objShell.Run(strWorkerCmd, 7, TRUE)
	errorcode = Err.Number
	On Error Goto 0
	
	If errorcode <> 0 Then
		BEnd_LogMessage "mebdd_worker shell object returns error code: " & errorcode
		BEnd_LastError = ""
		BEnd_CreateFileSystem = False
		Exit Function
	End If

	BEnd_LogMessage "mebdd_worker returns: " & exitcode
	
	If (exitcode = 255) Then
		BEnd_LastError = ""
		BEnd_CreateFilesystem = True
	Else
		BEnd_LastError = exitcode
		BEnd_CreateFilesystem = False
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
	Dim strLocalPath, strLocalBasename, strLocalPathname, strUnzipCmd, exitcode
	
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

	exitcode = objShell.Run(strUnzipCmd, 7, TRUE)

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

' works like the printf-function in C.
' takes a string with format characters and an array to expand.
'
' the format characters are always "%x", independ of the type.
'
' usage example:
'	dim str
'	str = BEnd_FMT( "hello, Mr. %x, today's date is %x.", Array("Miller",Date) )
'	response.Write str
'
' http://www.codeproject.com/Articles/250/printf-like-Format-Function-in-VBScript
Function BEnd_FMT( str, args )
	Dim res		' the result string.
	res = ""

	Dim pos		' the current position in the args array.
	pos = 0

	Dim i
	for i = 1 to Len(str)
		' found a fmt char.
		if Mid(str,i,1)="%" then
			if i<Len(str) then
				' normal percent.
				if Mid(str,i+1,1)="%" then
					res = res & "%"
					i = i + 1

				' expand from array.
				elseif Mid(str,i+1,1)="x" then
					res = res & CStr(args(pos))
					pos = pos+1
					i = i + 1
				end if
			end if

		' found a normal char.
		else
			res = res & Mid(str,i,1)
		end if
	next

	BEnd_FMT = res
End Function

' Wait for Time milliseconds
' http://stackoverflow.com/questions/16865430/sleep-routine-for-hta-scripts
Sub BEnd_Wait(Time)
	Dim wmiQuery, objWMIService, objPing, objStatus
	wmiQuery = "Select * From Win32_PingStatus Where Address = '1.1.1.1' AND Timeout = " & Time
	Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
	Set objPing = objWMIService.ExecQuery(wmiQuery)
	For Each objStatus in objPing
	Next
End Sub
