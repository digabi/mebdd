'    Copyright Matti Lattu 2014
'
'    This file is part of VBgettext.
'
'    VBgettext is free software: you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    any later version.
'
'    VBgettext is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License
'    along with VBgettext.  If not, see <http://www.gnu.org/licenses/>.
'
' Create PO template from .vbs or .hta file
'    Example:
'    cscript //Nologo example.hta >example.pot
'

Option Explicit

Dim arrCmdline, arrPathScript, objFSO, fsStderr
Set arrCmdline = Wscript.Arguments
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set fsStderr = objFSO.GetStandardStream (2)

Sub ProcessFile (strFilePath)
	Dim objFile, objRE, objREMatches, strLine, intLineCount, objThisMatch, objThisSubMatches
	
	Set objFile = CreateObject("Scripting.FileSystemObject").OpenTextFile(strFilePath,1)
	
	intLineCount = 0
	
	do while not objFile.AtEndOfStream
		strLine = objFile.ReadLine()
		intLineCount = intLineCount + 1
		
		' Look for VBGT_Get()
		Set objRE = new regexp
		
		objRE.Pattern = "VBGT_Get\(""(.+?)""\)"
		objRE.Global = True
		Set objREMatches = objRE.Execute(strLine)
		For Each objThisMatch in objREMatches
			Set objThisSubMatches = objThisMatch.SubMatches
			Wscript.Echo "#" & strFilePath & ":" & CStr(intLineCount)
			Wscript.Echo "msgid """ & objThisSubMatches.Item(0) & """"
			Wscript.Echo "msgstr """""
			Wscript.Echo ""
		Next
		
		' Look for VBGT_NGet()
		' This has not been tested and is not implemented in vbgettext.vbs
		Set objRE = new regexp
		objRE.Pattern = "VBGT_NGet\(""(.+?)"",""(.+?)"",.+\)"
		objRE.Global = True
		Set objREMatches = objRE.Execute(strLine)
		For Each objThisMatch in objREMatches
			Set objThisSubMatches = objThisMatch.SubMatches
			Wscript.Echo "#" & strFilePath & ":" & CStr(intLineCount)
			Wscript.Echo "# Warning: VBGT_NGet is not implemented in vbgettext.vbs"
			Wscript.Echo "msgid """ & objThisSubMatches.Item(0) & """"
			Wscript.Echo "msgid_plural """ & objThisSubMatches.Item(1) & """"
			Wscript.Echo "msgstr[0] """""
			Wscript.Echo "msgstr[1] """""
			Wscript.Echo ""
		Next
		
	loop
	
	objFile.Close
	Set objFile = Nothing
End Sub


Sub PrintUsage
	fsStderr.WriteLine "usage: cscript //Nologo vbgettext_createpo.vbs your.hta >your.pot"
	Wscript.Quit 1
End Sub


' If we don't have exactly one argument print usage
If arrCmdline.Count <> 1 Then
	fsStderr.WriteLine "Please give the vbs/hta file name as the first parameter."
	fsStderr.WriteLine ""
	PrintUsage
End If

' Make sure that the first argument points to an existing file
If not objFSO.FileExists(arrCmdline.item(0)) Then
	fsStderr.WriteLine "File '" & arrCmdline.item(0) & "' does not exist"
	fsStderr.WriteLine ""
	PrintUsage
End If

ProcessFile(arrCmdline.item(0))