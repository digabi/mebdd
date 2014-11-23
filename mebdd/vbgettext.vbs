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

Option Explicit

Dim VBGT_Data_Singular, VBGT_Data_Plural
Set VBGT_Data_Singular = CreateObject("Scripting.Dictionary")
Set VBGT_Data_Plural = CreateObject("Scripting.Dictionary")

Const VBGT_SINGULAR = 1
Const VBGT_PLURAL = 2

' Load PO file to VBGT_Data_Singular, VBGT_Data_Plural
' strFilePath should point to .po file
' strFilePath "-" removes all existing data and uses default msgid:s
'
' PO file format:
' https://www.gnu.org/software/gettext/manual/html_node/PO-Files.html
' Returns true if PO file was loaded OK. Otherwise false.
Function VBGT_GetPO(strFilePath)
	Dim objFSO, objFile, strLine, strCmd, strParam, strMsgidSingular, strMsgidPlural, intCmdIndex
	
	' Remove all existing data
	VBGT_Data_Singular.RemoveAll()
	VBGT_Data_Plural.RemoveAll()
	
	' If strFilePath is "-" the user wishes to use the default language
	If strFilePath = "-" Then
		VBGT_GetPO = True
		Exit Function
	End If
	
	' If given strFilePath does not point to an existing file return false
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	If not objFSO.FileExists(strFilePath) Then
		VBGT_GetPO = False
		Exit Function
	End If
	
	Set objFile = CreateObject("Scripting.FileSystemObject").OpenTextFile(strFilePath,1)
	do while not objFile.AtEndOfStream
		strLine = objFile.ReadLine()
		
		If InStr(strLine, " ") > 0 Then
			strCmd = LCase(Left(strLine, InStr(strLine, " ")-1))
			strParam = Mid(strLine, InStr(strLine, " ")+1)
			
			' Remove surrounding quotation marks from strParam
			If Left(strParam, 1) = """" and Right(strParam, 1) = """" Then
				strParam = Mid(strParam, 2, Len(strParam)-2)
			End If
			
			' Get possible intCmdIndex
			If InStr(strCmd, "[") > 0 and Right(strCmd, 1) = "]" Then
				intCmdIndex = CInt(Mid(strCmd, InStr(strCmd, "[")+1, Len(strCmd)-Instr(strCmd, "[")-1))
			End If
			
			If strCmd = "msgid" Then
				strMsgidSingular = strParam
				strMsgidPlural = ""
				intCmdIndex = -1
			End If
			
			If strCmd = "msgid_plural" Then
				strMsgidPlural = strParam
			End If
			
			'MsgBox "line: " & strLine & " msgid: " & strMsgidSingular & " msgid_plural: " & strMsgidPlural & " intCmdIndex: " & CStr(intCmdIndex)
			
			If strCmd = "msgstr" Then
				' This specifies singular value
				If IsEmpty(strMsgidSingular) Then
					' Error: Met msgstr but msgid is empty
					Err.Raise 500
				Else
					' Add singular value only
					VBGT_Add_Singular strMsgidSingular,strParam
				End If
			End If
			
			If Left(strCmd, 7) = "msgstr[" Then
				' This specifies plural value
				If IsEmpty(strMsgidSingular) Then
					' Error: Met msgstr but msgid is empty
					Err.Raise 500
				Else
					' Add singular value
					VBGT_Add_Singular strMsgidSingular,strParam
				End If
				
				If IsEmpty(strMsgidPlural) Then
					' Error: Met msgstr[] but msgid_plural is empty
					Err.Raise 500
				Else
					' Add plural value
					VBGT_Add_Plural strMsgidPlural,strParam,intCmdIndex
				End If
			End If
			
		End If
	Loop

	objFile.Close
	Set objFile = Nothing
	
	VBGT_GetPO = True
End Function

' Add singular value to VBGT_Data_Singular
Sub VBGT_Add_Singular(strMsgid, strMsgstr)
	If strMsgid = "" or strMsgstr = "" Then
		Exit Sub
	End If
	
	If VBGT_Data_Singular.Exists(strMsgid) Then
		' Change the existing value
		VBGT_Data_Singular(strMsgid) = strMsgstr
	Else
		' No previous value exists
		VBGT_Data_Singular.Add strMsgid, strMsgstr
	End If
End Sub

' Add plural value to VBGT_Data_Plural
Sub VBGT_Add_Plural(strMsgid, strMsgstr, intIndex)
	If strMsgid = "" or strMsgstr = "" Then
		Exit Sub
	End If

	Dim objSD
	
	If not VBGT_Data_Plural.Exists(strMsgid) Then
		' This is the first value, create array
		Set objSD = CreateObject("Scripting.Dictionary")
		VBGT_Data_Plural.Add strMsgid,objSD
	End If
		
	VBGT_Data_Plural(strMsgid)(intIndex) = strMsgstr
End Sub

' Equals to dgettext(): returns the translated string for the given strMSGID
Function VBGT_Get(strMSGID)
	If VBGT_Data_Singular.Exists(strMSGID) Then
		VBGT_Get = VBGT_Data_Singular(strMSGID)
	Else
		VBGT_Get = strMSGID
	End If
End Function

' Equals to dngettext(): returns the translated string for the given strMSGID or strMSGID_PLURAL
' depending on intN.

Function VBGT_NGet(strMSGID, strMSGID_PLURAL, intN)
	' FIXME
End Function
