<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" %>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
	products_sorted = Request.Form("products")
	DataConn.BeginTrans 
	Set objCmd = Server.CreateObject ("ADODB.Command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "DELETE FROM TBL_New_Products_Sortable; INSERT TBL_New_Products_Sortable (product_id, custom_sorting) VALUES " & products_sorted 
	Response.Write objCmd.CommandText
	objCmd.Execute()
  
	If Err.Number = 0 Then  
		DataConn.CommitTrans  
	Else  
		DataConn.RollbackTrans  
	End If 

	DataConn.Close()
%>