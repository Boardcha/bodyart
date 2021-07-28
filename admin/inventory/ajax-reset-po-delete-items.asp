<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" %>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
	' Delete all the purchase order items that are currently set to a temp # for that brand	
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "DELETE FROM tbl_po_details WHERE po_orderid = 0 AND po_temp_id = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("po_temp_id",3,1,10,request("tempid")))
	objCmd.Execute()	
	
	response.cookies("po-filter-status") = ""
	response.cookies("po-filter-active") = ""
	response.cookies("po-filter-qty") = ""


DataConn.Close()
%>