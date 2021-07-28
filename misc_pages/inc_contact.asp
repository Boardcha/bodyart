<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/template/inc_includes.asp" -->
<%
If Request.Form("invoice") <> "" AND (Request.Form("reason") = "Order - Change/Update" OR Request.Form("reason") = "Order - Problem") then

	var_invoice_num = Request.Form("invoice")

If IsNumeric(var_invoice_num) Then
	' leave var_invoice_num as is
Else
	' reset to 0, prevents asp errors from random entries in querystring
	var_invoice_num = 0
End If

	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "UPDATE sent_items SET shipped = 'ON HOLD' WHERE (shipped = 'Pending...' OR shipped = 'Review') AND ID = ? AND email = ?"
	objCmd.Prepared = true
	objCmd.Parameters.Append(objCmd.CreateParameter("ID",3,1,12,var_invoice_num))
	objCmd.Parameters.Append(objCmd.CreateParameter("email",200,1,50,Request.form("email")))
	Set rsUpdateInvoice = objCmd.Execute()

End if

mailer_type = "contact-us"
%>
<!--#include virtual="/emails/function-send-email.asp"-->
<!--#include virtual="/emails/email_variables.asp"-->