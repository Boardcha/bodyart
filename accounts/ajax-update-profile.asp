<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/template/inc_includes_ajax.asp" -->
<%
' Get customer info from database
set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT * FROM customers WHERE customer_ID = ?"
objCmd.Parameters.Append(objCmd.CreateParameter("CustID_Cookie",3,1,10,CustID_Cookie))
Set rsGetUser = objCmd.Execute()

 
  		' Update email in BAF database
		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "UPDATE customers SET customer_first = ?, customer_last = ? WHERE customer_ID = ?"
		objCmd.Parameters.Append(objCmd.CreateParameter("first",200,1,50, request.form("first")))
		objCmd.Parameters.Append(objCmd.CreateParameter("last",200,1,50, request.form("last")))
		objCmd.Parameters.Append(objCmd.CreateParameter("CustomerID",3,1,10, CustID_Cookie))
		objCmd.Execute()
	

DataConn.Close()
Set DataConn = Nothing
%>