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
If isEmpty(KT_CoreFunctions__ALREADYLOADED) Then
	KT_CoreFunctions__ALREADYLOADED = True
	KT_uploadErrorMsg = "<strong>File not found:</strong> <br />###<br /><strong>Please upload the includes/ folder to the testing server.</strong>"
	Sub KT_createGlobalFSO()
		If Not isObject(KT_FSO) Then
			ExecuteGlobal "Set KT_FSO = Server.CreateObject(""Scripting.FileSystemObject"")"	
		End If	
	End Sub
	Function KT_SetPathSessions()
		KT_createGlobalFSO
		' sets 2 sessions: KT_AbsolutePathToRootFolder and KT_SiteURL
		
		' In order to know which one is the root folder, must check for (a) particular folder(s) 
		' that we know for sure that is(are) located in the root folder of the site
		Dim SearchForFolderName: SearchForFolderName = "includes\common" 

		Dim url: url = Request.ServerVariables("URL")
		' cut the trailing /
		LastSeparator = InStrRev(url, "/")
		If LastSeparator > 0 Then
			url = left(url, LastSeparator-1)
		End If

		Dim path: path = Server.MapPath(".") & "\"
		' cut the trailing \
		LastSeparator = InStrRev(path, "\")
		If LastSeparator > 0 Then
			path = left(path, LastSeparator-1)
		End If
		
		Dim prefix: prefix = ""
		Dim found: found = False
		Do while not Found
			If KT_FSO.FolderExists (path &  "\" & SearchForFolderName) Then
				Found = true
				Exit do
			Else
				' remove one folder lever both from path and url
				LastSeparator = InStrRev(url, "/")
				If LastSeparator > 0 Then
					url = left(url, LastSeparator-1)
				Else
					found = true ' force the exit from loop	
				End If
		
				LastSeparator = InStrRev(path, "\")
				If LastSeparator > 0 Then
					path = left(path, LastSeparator-1)
					prefix = prefix & "..\"
				Else
					found = true ' force the exit from loop	
				End If				 	
			End If						
		Loop

		If prefix = "" Then
			prefix = "."
		End If
		If found Then
			Session("KT_SitePath") = url
			Session("KT_AbsolutePathToRootFolder") = KT_FSO.GetAbsolutePathName(Server.MapPath(".") & "\" & prefix) & "\"
		Else
			KT_GetAbsolutePathToRootFolder = ""
		End If	
	End Function

	' retrieves the path on disk to the site root (eg C:\www\sites\MYSITE\)
	Function KT_GetAbsolutePathToRootFolder()
		If Session("KT_AbsolutePathToRootFolder") = ""  Then
			KT_SetPathSessions
		End If
		KT_GetAbsolutePathToRootFolder	= Session("KT_AbsolutePathToRootFolder")
	End Function
	

	Sub KT_LoadASPFiles (arrPathsRelativeToRoot)
		KT_createGlobalFSO
		absolutePathToRootFolder = KT_GetAbsolutePathToRootFolder()
		
		On Error Resume Next	
		Dim i
		For i=0 to ubound(arrPathsRelativeToRoot)		
			pathRelativeToRootFolder = arrPathsRelativeToRoot(i)
			absolutePathToFile = absolutePathToRootFolder & replace(pathRelativeToRootFolder, "/", "\")
			If KT_FSO.FileExists(absolutePathToFile) Then
				' read the file content
				Dim f: Set f = KT_FSO.OpenTextFile(absolutePathToFile, 1, False)
				content = f.ReadAll
				f.Close
				Set f = nothing
				
				If instr(pathRelativeToRootFolder, "KT_common.asp") <> 0 Then
					content = mid(content, 1, instr(content, "<S" & "CRIPT language=""jscript"" runat=""server")-1)
				End If
				If instr(pathRelativeToRootFolder, "tNG.inc.asp") <> 0 Then
					content = mid(content, 1, instr(content, "<S" & "CRIPT language=""jscript"" runat=""server")-1)
				End If								
				' replace ASP tags 
				execcontent = replace (content, "<" & "%", "")
				execcontent = replace (execcontent, "%" & ">", "")
				ExecuteGlobal(execcontent)
			Else
				Session.Contents.RemoveAll
				Response.write replace(KT_uploadErrorMsg, "###", pathRelativeToRootFolder)
				Response.End()
			End If
			If err.number<>0 Then
				Response.write "<br><span style=""color:red"">Error loading file '" & pathRelativeToRootFolder & "'<br>" & err.description & "</font>"
				Response.End()
			End If
		Next
		On Error GoTo 0
	End Sub

End If


	' load KT_common
	If isEmpty(KT_Common__ALREADYLOADED) Then
		KT_LoadASPFiles Array("includes/common/KT_common.asp")
	End If	
	
	If isEmpty(KT_ResourcesFunctions__ALREADYLOADED) Then
		KT_ResourcesFunctions__ALREADYLOADED = True
		KT_LoadASPFiles Array("includes/common/lib/resources/KT_ResourcesFunctions.inc.asp")
	End If
%>
