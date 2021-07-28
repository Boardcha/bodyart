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
	' create interakt object if not already created
	If not isObject(interakt) Then
		ExecuteGlobal ("Dim interakt" & vbNewLine & "Set interakt = Server.CreateObject(""Scripting.Dictionary"")")
	End If	
	
	If not isObject(interakt("resources")) Then
		 Set interakt("resources") = Server.CreateObject("Scripting.Dictionary")
	End If


	Function KT_getResource(resourceName, dictionary, args)
		If isnull(resourceName) Then
			resourceName = "default"
		End If
		If isnull(dictionary) Then
			dictionary = "default"
		End If
		If isnull(args) Or isempty(args) Or (Not isarray(args)) Then
			args = array()
		End If
		
		Dim resourceValue: resourceValue = resourceName
		Dim dictionaryFileName: 
			dictionaryFileName = KT_GetAbsolutePathToRootFolder() & "includes\resources\" & dictionary & ".res.asp"



		' First thing: check the dictionary for the corresponding resourceName
		If Not isObject(interakt("resources")(dictionary)) Then
			' must load the dictionary
			Dim fso: Set fso = Server.CreateObject("Scripting.FileSystemObject")
			If fso.FileExists(dictionaryFileName) Then
				' read the file content
				KT_oriental_language = false
				If Not KT_oriental_language Then
					Dim f: Set f = fso.OpenTextFile(absolutePathToResourceDictionaries & dictionaryFileName, 1, False)
					content = f.ReadAll
					f.Close
					Set f = nothing
				Else
					Set streamFile = Server.CreateObject("ADODB.Stream")
					streamFile.Type = 2
					streamFile.Charset = "Shift_JIS" ' change the charset accordingly; defaults to japanesse
					streamFile.Open
					streamFile.LoadFromFile absolutePathToResourceDictionaries & dictionaryFileName
					content = streamFile.readText(-1)
					streamFile.Close                    
					Set streamFile = nothing
				End If
					
				execcontent = replace (content, "<" & "%", "")
				execcontent = replace (execcontent, "%" & ">", "")
				Execute execcontent
				
				If isObject(res) Then
					Set interakt("resources")(dictionary) = res
				End If	
			End If
			Set fso = nothing

                        dictionaryFileName = KT_GetAbsolutePathToRootFolder() & "includes\resources\" & dictionary & "_pro.res.asp"
                        Set fso = Server.CreateObject("Scripting.FileSystemObject")
			If fso.FileExists(dictionaryFileName) Then
				' read the file content
				KT_oriental_language = false
				If Not KT_oriental_language Then
					Set f = fso.OpenTextFile(absolutePathToResourceDictionaries & dictionaryFileName, 1, False)
					content = f.ReadAll
					f.Close
					Set f = nothing
				Else
					Set streamFile = Server.CreateObject("ADODB.Stream")
					streamFile.Type = 2
					streamFile.Charset = "Shift_JIS" ' change the charset accordingly; defaults to japanesse
					streamFile.Open
					streamFile.LoadFromFile absolutePathToResourceDictionaries & dictionaryFileName
					content = streamFile.readText(-1)
					streamFile.Close                    
					Set streamFile = nothing
				End If
					
				execcontent = replace (content, "<" & "%", "")
				execcontent = replace (execcontent, "%" & ">", "")
				Execute execcontent
				
				If isObject(res) Then
                                  For each key in res
                                    interakt("resources")(dictionary)(key) = res(key)
                                  Next
				End If	
			End If
			Set fso = nothing
		End If

		foundResource = false
		If isObject(interakt("resources")(dictionary)) Then
			If interakt("resources")(dictionary).Exists(resourceName) Then
				foundResource = true
				resourceValue = interakt("resources")(dictionary)(resourceName)
			End If
		End IF	
		
		If Not foundResource Then
			'If trim(resourceName) <> "" And trim(resourceName) <> "%s" Then
			'	Response.write "<br />Resource '" & resourceName & "' not defined in dictionary '" & dictionary & "'.<br />"
			'	Response.End()
			'End If

			If right(resourceValue, 2)= "_D" Then
				resourceValue = left(resourceValue, len(resourceValue)-2)
			End If
		End If
				
		If ubound(args) <> -1 Then
			resourceValue = KT_sprintf(resourceValue, args)
		End If
		
		KT_getResource = resourceValue
	End Function
%>