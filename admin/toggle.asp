<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
toggleItem = Request.Form("toggleItem")
isChecked = Request.Form("isChecked")

If isChecked = "true" OR isChecked = "false" Then
	set objCmd = Server.CreateObject("ADODB.Command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "UPDATE TBL_Toggle_Items SET value=? WHERE toggle_item = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("value", 11, 1, 10, isChecked))
	objCmd.Parameters.Append(objCmd.CreateParameter("toggle_item", 200, 1, 50, toggleItem))
	objCmd.Execute()
End If

%>