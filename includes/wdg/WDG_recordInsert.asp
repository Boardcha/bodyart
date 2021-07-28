<!-- #include file="WDG.asp" -->
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
	Function printErrorScript(desc)
		Response.Write("<html><body onLoad=""parent.MXW_DynamicObject_reportDone('" & Request.QueryString("el") & "', true, '" & KT_escapeJS(desc) & "')""></body></html>")
		Response.End()
	End Function
	
	id = KT_getRealValue("GET", "id")
	el = KT_getRealValue("GET", "el")
	text =  KT_getRealValue("GET", "text")

	If isnull(id) Or id="" OR isnull(el) Or el = "" Then Response.End()
	If TypeName(Session("WDG_sessInsTest")) <> "Dictionary" Then Response.End()

	Dim WDG_sessInsTest, vars
	Set WDG_sessInsTest = Session("WDG_sessInsTest")

	id = CInt(id)
	If NOT KT_isset(WDG_sessInsTest(id)) Then Response.End()

	Set vars = WDG_sessInsTest(id)
	KT_LoadASPFiles Array("Connections\" & vars("conn") & ".asp")

	On Error Resume Next

	Dim rs, Cnxn 
	Set Cnxn = Server.CreateObject("ADODB.Connection")
	ExecuteGlobal("strCnxn = MM_" & vars("conn") & "_STRING")
	Cnxn.Open strCnxn
	
	KT_setDbType(Cnxn)
	
	If Err Then
		printErrorScript(Err.Description)
	End If

	sql = "insert into " & vars("table") & " (" & KT_escapeFieldName(vars("updatefield")) & ") values (" & KT_escapeForSql(text, "STRING_TYPE" ) & ") "
	Cnxn.Execute sql

	If Err Then
		printErrorScript(Err.Description)
	End If

	sql = "select " & KT_escapeFieldName(vars("idfield")) & " as id FROM " & vars("table") & " where " & KT_escapeFieldName(vars("updatefield")) & " = " & KT_escapeForSql(text, "STRING_TYPE")
	Set rs=Cnxn.Execute(sql)

	If Err Then
		printErrorScript(Err.Description)
	End If

	newid=rs("id")
	text=KT_escapeJS(text)
%>
<html><body onLoad="parent.MXW_DynamicObject_reportDone('<%=el%>', isError)">
<script>
	var isError = false;
	var targetRSName = '<%= vars("rsName")%>';
	var targetEditableDropdownName = '<%=el%>';
	var idfield = '<%=vars("idfield")%>';
	var updatefield = '<%=vars("updatefield")%>';
	var insertedID = '<%=newid%>';
	var insertedValue = '<%=text%>';

	for(dyninputname in parent[parent.$DYS_GLOBALOBJECT]) {
		updatedDynamicInput = parent[parent.$DYS_GLOBALOBJECT][dyninputname]
		if (!updatedDynamicInput || updatedDynamicInput && !updatedDynamicInput.oldinput) {
			continue;
		}
		
		if(updatedDynamicInput.edittype != 'E') {
			continue;
		}
		
		recordsetName = parent.WDG_getAttributeNS(updatedDynamicInput.oldinput, 'recordset');
		if (targetRSName != recordsetName) {
			continue;
		}
	
		if (targetEditableDropdownName == dyninputname) {
			var newRow = [];
			newRow[idfield] = insertedID;
			newRow[updatefield] = insertedValue;
			updatedDynamicInput.recordset.Insert(newRow, parseInt(updatedDynamicInput._firstMatch, 10) + 1);

			updatedDynamicInput.oldinput.options.add(new parent.Option(insertedValue, insertedID));
			updatedDynamicInput.sel.options.add(new parent.Option(insertedValue, insertedID));
			updatedDynamicInput.addButton.disabled = true;
			updatedDynamicInput.oldinput.selectedIndex = updatedDynamicInput.oldinput.options.length - 1;
			updatedDynamicInput.sel.selectedIndex = updatedDynamicInput.sel.options.length - 1;
			updatedDynamicInput.oldinput.value = insertedID;
			updatedDynamicInput.newvalue = insertedID;
			parent.MXW_DynamicObject_syncSelection(dyninputname, false, true);
			updatedDynamicInput.edit.focus();	
		} else {
			updatedDynamicInput.oldinput.options.add(new parent.Option(insertedValue, insertedID));
			updatedDynamicInput.sel.options.add(new parent.Option(insertedValue, insertedID));
		}
	}
	var isComplete = true;
</script>
</body>
</html>
