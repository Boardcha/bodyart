<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" %>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
if request.form("column") = "sku" then

	Set objCmd = Server.CreateObject ("ADODB.Command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "UPDATE ProductDetails SET detail_code = ? WHERE ProductID = ?" 
	objCmd.Parameters.Append(objCmd.CreateParameter("productid",200,1,50, request.form("column_value")))
	objCmd.Parameters.Append(objCmd.CreateParameter("productid",3,1,15,request.form("productid")))
	Set rs_GetImgID = objCmd.Execute()

end if
DataConn.Close()
%>