<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/template/inc_includes_ajax.asp" -->
<%

set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT ID, amount, code, invoice FROM TBLcredits WHERE code = ? AND amount <> 0"
objCmd.Parameters.Append(objCmd.CreateParameter("CertCode",200,1,30, request.form("cert_code")))
Set rsGetCertificate = objCmd.Execute()

if not rsGetCertificate.EOF then


	' Add gift cert amount to customer store credit
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "UPDATE customers SET credits = ? + credits WHERE customer_ID = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("Credits",6,1,10, rsGetCertificate.Fields.Item("amount").Value))
	objCmd.Parameters.Append(objCmd.CreateParameter("CustomerID",3,1,12,CustID_Cookie))
	objCmd.Execute()

	' Set gift certificate to $0 balance
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "UPDATE TBLCredits SET amount = 0, custid_converted = ? WHERE ID = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("CustomerID",3,1,12,CustID_Cookie))
	objCmd.Parameters.Append(objCmd.CreateParameter("CertCode",200,1,30, rsGetCertificate.Fields.Item("ID").Value))
	objCmd.Execute()
	
	' add a note to the order about the gift certificate conversion
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "INSERT INTO tbl_invoice_notes (user_id, invoice_id, note) VALUES (?,?,?)"
	objCmd.Parameters.Append(objCmd.CreateParameter("user_id",3,1,10,28))
	
	if rsGetCertificate.Fields.Item("invoice").Value <> "" then
		var_invoice = rsGetCertificate.Fields.Item("invoice").Value
	else
		var_invoice = 0
	end if
	
	objCmd.Parameters.Append(objCmd.CreateParameter("invoiceid",3,1,10,var_invoice))
	objCmd.Parameters.Append(objCmd.CreateParameter("note",200,1,250,"Customer converted gift certificate to store credit via the website on " & now()))
	objCmd.Execute()
	
	' Get current store credit amount after gift cert has been added to it
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT credits FROM customers WHERE customer_ID = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("CustID_Cookie",3,1,10,CustID_Cookie))
	Set rsNewAmount = objCmd.Execute()
	
	
%>
{
	"status":"success",
	"amount":"<%= FormatCurrency(rsNewAmount.Fields.Item("credits").Value, 2) %>"
}
<%
else
%>
{
	"status":"fail"
}
<%
end if

DataConn.Close()
Set DataConn = Nothing
%>
