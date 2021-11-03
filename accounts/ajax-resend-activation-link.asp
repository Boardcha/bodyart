<% @LANGUAGE="VBSCRIPT" %>
<!--#include virtual="/template/inc_includes.asp" -->
<%
mailer_type = "account activation"

set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT email, activation_hash FROM customers WHERE customer_ID = ?"
objCmd.Parameters.Append(objCmd.CreateParameter("CustID_Cookie",3,1,10,CustID_Cookie))
Set rsGetActivationInfo = objCmd.Execute()

if NOT rsGetActivationInfo.EOF then
    email = rsGetActivationInfo("email")
    activation_hash = rsGetActivationInfo("activation_hash")			
%>
	<!--#include virtual="emails/function-send-email.asp"-->
	<!--#include virtual="emails/email_variables.asp"-->
<%
end if

DataConn.Close()
Set DataConn = Nothing
%>