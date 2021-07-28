<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" %>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT seo_meta_title FROM jewelry WHERE seo_meta_title = ? AND productid <> ?"
    objCmd.Parameters.Append(objCmd.CreateParameter("seo_title",200,1,100, request.form("seo_title")))
    objCmd.Parameters.Append(objCmd.CreateParameter("productid",3,1,15, request.form("id")))
	set rsCheckDupe = objCmd.Execute()
	
if rsCheckDupe.eof then
%>
{
    "status":"success"    
}
<% else %>
{
    "status":"fail"    
}
<% end if

Set rsCheckDupe = nothing
DataConn.Close()
%>