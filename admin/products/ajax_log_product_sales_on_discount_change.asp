<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" %>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
	product_id = Request.Form("product_id")
	discount = Request.Form("discount")
	
	Set objCmd = Server.CreateObject ("ADODB.Command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandType = 4 'Stored Procedure
	objCmd.CommandText = "SP_Log_Sales_On_Discount_Change"
	objCmd.Parameters.Append(objCmd.CreateParameter("@product_id",3,1,10, product_id))
	objCmd.Parameters.Append(objCmd.CreateParameter("@discount",3,1,10, discount))
	Response.Write objCmd.CommandText
	objCmd.Execute()
	
 	DataConn.Close()
%>