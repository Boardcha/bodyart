<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
autoclave = Request.Form("autoclave")
If autoclave = 0 OR autoclave = 1 Then
	set objCmd = Server.CreateObject("ADODB.Command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "UPDATE TBL_Toggle_Items SET value=? WHERE toggle_item='toggle_autoclave'"
	objCmd.Parameters.Append(objCmd.CreateParameter("autoclave", 3, 1, 15, autoclave))
	objCmd.Execute()
End If
%>