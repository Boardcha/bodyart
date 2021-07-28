<%
'
' ADOBE SYSTEMS INCORPORATED
' Copyright 2007 Adobe Systems Incorporated
' All Rights Reserved
' 
' NOTICE:  Adobe permits you to use, modify, and distribute this file in accordance with the 
' terms of the Adobe license agreement accompanying it. If you have received this file from a 
' source other than Adobe, then your use, modification, or distribution of it requires the prior 
' written permission of Adobe.
'

'
'	Copyright (c) InterAKT Online 2000-2005
'
Function WDG_aspdict2jsarray(ByRef obj)
	tm = ""
	If obj.Count = 0 Then
		tm = "[]"
	Else
		tm = tm & "["
		For each v in obj
			If isnull(obj(v)) Then
				tm = tm & "'" & "'" &  "," 			
			Else
				tm = tm & "'" & KT_escapeJS(obj(v)) & "'" &  "," 
			End If	
		Next
		tm = left (tm, len(tm) - 1)
		tm = tm  & "]"
	End If
	WDG_aspdict2jsarray = tm
End Function

Function WDG_registerRecordInsert(connectionName, rsName, idField, updateField) 
	If NOT KT_isSet(KT_dynamicInputSW) Then
		ExecuteGlobal("Dim KT_dynamicInputSW: KT_dynamicInputSW = true")
		Dim WDG_sessInsTest
		Set WDG_sessInsTest = Server.CreateObject("Scripting.Dictionary")
		Set Session("WDG_sessInsTest") = WDG_sessInsTest
	End If

	Execute("Dim rs: Set rs = " & rsName)

	'the sql query (string)
	sql_query = rs.Source
	Dim matches
	KT_preg_match "from\s*([^\s]+)", sql_query, matches
	sqlTable = KT_preg_replace("(from\s*)([^\s]+)", "$2", matches(0))

	Set WDG_sessInsTest = Session("WDG_sessInsTest")
	max_idx = -1
	For Each idx in WDG_sessInsTest
		If idx > max_idx Then
			max_idx = idx
		End If
	Next

	Set new_item = Server.CreateObject("Scripting.Dictionary")
	new_item("conn") = connectionName
	new_item("rsName") = rsName
	new_item("table") = sqlTable
	new_item("idfield") = idField
	new_item("updatefield") = updateField

	'the parameters required for the editable dropdown record insertion are stored in a Dictionary object in session
	'not an array as in previous versions or in PHP
	Set WDG_sessInsTest(max_idx + 1) = new_item
	Set Session("WDG_sessInsTest") = WDG_sessInsTest
End Function
%>