<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" %>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
	'========= DELETE PURCHASE ORDER ====================================
	set objCmd = Server.CreateObject("ADODB.Command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "DELETE FROM TBLWaitingList WHERE ID = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("waiting_list_id",3,1,15, request.form("id")))
	objCmd.Execute()
	
%>