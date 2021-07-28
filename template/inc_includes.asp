<!--#include virtual="/Connections/sql_connection.asp" -->
<!--#include virtual="/functions/sha256.asp"-->
<!--#include virtual="/functions/salt.asp"-->
<%

	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT * FROM customers WHERE customer_ID = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("CustID_Cookie",3,1,10, CustID_Cookie))
	Set rsGetUser = objCmd.Execute()
	
	if request.cookies("ip-country") <> "" then
		strcountryName = request.cookies("ip-country")
	else
		strcountryName = "US"
	end if
%>
