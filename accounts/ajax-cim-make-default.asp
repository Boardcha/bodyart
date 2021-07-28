<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/template/inc_includes_ajax.asp" -->
<%

' Make address default
If request.form("type") = "shipping" Then

	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "UPDATE TBL_AddressBook SET default_shipping = 0 WHERE custID = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("cim_custid",200,1,30,CustID_Cookie))
	objCmd.Execute()

	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "UPDATE TBL_AddressBook SET default_shipping = 1 WHERE custID = ? AND ID = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("cim_custid",200,1,30,CustID_Cookie))
	objCmd.Parameters.Append(objCmd.CreateParameter("shipping_id",3,1,10,request.form("id")))
	objCmd.Execute()
	
end if


If request.form("type") = "billing" Then

	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "UPDATE TBL_AddressBook SET default_billing = 0 WHERE custID = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("cim_custid",200,1,30,CustID_Cookie))
	objCmd.Execute()
	
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "UPDATE TBL_AddressBook SET default_billing = 1 WHERE custID = ? AND ID = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("cim_custid",200,1,30,CustID_Cookie))
	objCmd.Parameters.Append(objCmd.CreateParameter("shipping_id",3,1,10,request.form("id")))
	objCmd.Execute()
		
End If




DataConn.Close()
Set DataConn = Nothing
%>
