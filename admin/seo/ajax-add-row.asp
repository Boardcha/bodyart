<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" %>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<%

	set objCmd = Server.CreateObject("ADODB.Command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "INSERT INTO tbl_sitemap_searches DEFAULT VALUES"
    objCmd.Execute()
    
    set objCmd = Server.CreateObject("ADODB.Command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT TOP(1) id FROM tbl_sitemap_searches ORDER BY id DESC"
    set rsGetID = objCmd.Execute()
    
if NOT rsGetID.eof then
%>
{
    "id":"<%= rsGetID.Fields.Item("id").Value %>",
    "status":"success"
}
<% 
else
%>
{
    "status":"fail"
}
<%
end if

DataConn.Close()
%>