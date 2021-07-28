<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" %>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
' get product information
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT picture, title FROM jewelry WHERE ProductID = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("productid",3,1,10,request.form("productid")))
	Set getProduct = objCmd.Execute()
	
if not getProduct.eof then
%>
<img src="http://bafthumbs-400.bodyartforms.com/<%= getProduct.Fields.Item("picture").Value %>" width="150" height="150">

<%
end if


DataConn.Close()
Set rsResearch = Nothing
%>