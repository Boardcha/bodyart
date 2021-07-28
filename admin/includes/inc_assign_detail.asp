<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" %>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
if request.form("detailid") <> "" then ' check to see if a detailid is provided and if not, just update the main products table	

	' Activate detail, update qty, and re-assign bin
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "UPDATE ProductDetails SET qty = ?, BinNumber_Detail = ?, active = 1 WHERE ProductDetailID = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("qty",3,1,10,request.form("qty")))
	objCmd.Parameters.Append(objCmd.CreateParameter("bin",3,1,10,request.form("bin")))
	objCmd.Parameters.Append(objCmd.CreateParameter("detailid",3,1,10,request.form("detailid")))
	Set rsUpdate = objCmd.Execute()
	
	' Activate main product
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "UPDATE jewelry SET jewelry.active = 1 FROM jewelry INNER JOIN ProductDetails ON jewelry.ProductID = ProductDetails.ProductID WHERE ProductDetailID = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("detailid",3,1,10,request.form("detailid")))
	Set rsUpdate = objCmd.Execute()

end if

DataConn.Close()
Set rsResearch = Nothing
%>