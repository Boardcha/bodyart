<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" %>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
if request.form("detailid") <> "" then ' check to see if a detailid is provided and if not, just update the main products table	
	
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "update ProductDetails set qty= ? where ProductDetailID = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("qty",3,1,10,request.form("qty")))
	objCmd.Parameters.Append(objCmd.CreateParameter("detailid",3,1,10,request.form("detailid")))
	Set rsUpdate = objCmd.Execute()

end if

DataConn.Close()
Set rsResearch = Nothing
%>