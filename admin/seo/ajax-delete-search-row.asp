<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" %>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<%

	set objCmd = Server.CreateObject("ADODB.Command")
	objCmd.ActiveConnection = DataConn
    objCmd.CommandText = "DELETE FROM tbl_sitemap_searches WHERE id = ?"
    objCmd.Parameters.Append(objCmd.CreateParameter("id",200,1,30,request.form("id")))
    objCmd.Execute()
    
DataConn.Close()
%>