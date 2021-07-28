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
	If IsEmpty(KT_ResourcesFunctions__ALREADYLOADED) Then
		KT_LoadASPFiles Array("includes/common/lib/resources/KT_Resources.asp")
	End If	

	If isEmpty(KT_FakeRecordset__ALREADYLOADED) Then
		KT_FakeRecordset__ALREADYLOADED = True
		KT_LoadASPFiles Array("includes/common/lib/db/KT_FakeRecordset.class.asp")
	End If
%>