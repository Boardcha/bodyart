<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
Dim GetNotes__MMColParam
GetNotes__MMColParam = "1"
If (Request.Form("InvoiceID") <> "") Then 
  GetNotes__MMColParam = Request.Form("InvoiceID")
End If
%>
<%
Dim GetNotes
Dim GetNotes_cmd
Dim GetNotes_numRows

Set GetNotes_cmd = Server.CreateObject ("ADODB.Command")
GetNotes_cmd.ActiveConnection = MM_bodyartforms_sql_STRING
GetNotes_cmd.CommandText = "SELECT ID, email, country FROM dbo.sent_items WHERE ID = ?" 
GetNotes_cmd.Prepared = true
GetNotes_cmd.Parameters.Append GetNotes_cmd.CreateParameter("param1", 5, 1, -1, GetNotes__MMColParam) ' adDouble

Set GetNotes = GetNotes_cmd.Execute
GetNotes_numRows = 0
%>
<% if (GetNotes.Fields.Item("country").Value) <> "USA" then
shipping = "DHL Global basic ground"
else
shipping = "DHL Basic mail"
end if
%>
<% if request.form("status") = "ORDER PROBLEM" then
OrderError = 0
else
OrderError = 1
end if 
%>
<% 
set CopyRow = Server.CreateObject("ADODB.Command")
CopyRow.ActiveConnection = MM_bodyartforms_sql_STRING
CopyRow.CommandText = "INSERT INTO sent_items (shipped, customer_ID, customer_first, customer_last, company, address, address2, city, state, province, zip, country, email, date_order_placed, shipping_rate, item_description, ship_code, phone, OrderError, AmountLost, Comments_OrderError, date_sent, our_notes, cc_num2, cc_month, cc_year, transactionID, pay_method, PackagedBy, ErrorType, UPS_Service, shipping_type) SELECT 'Pending...', customer_ID, customer_first, customer_last, company, address, address2, city, state, province, zip, country, email, '" & now() & "', 0, '<b><font size=3>REPLACEMENT ORDER</font></B><br>', 'paid', phone, "& OrderError & ", " & Request.Form("Lost") & ", Comments_OrderError, date_sent, '" & Request.Form("comments") & "', cc_num2, cc_month, cc_year, transactionID, pay_method, PackagedBy, '" & Request.Form("ErrorType") & "', '', '" & shipping & "' FROM sent_items WHERE ID =" & Request.Form("InvoiceID") 
CopyRow.Execute() 
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Create a new order to ship out</title>
<link href="../includes/nav.css" rel="stylesheet" type="text/css" />
</head>

<body class="mainbkgd">
<%
Dim rsGetNewInvoice
Dim rsGetNewInvoice_cmd
Dim rsGetNewInvoice_numRows

Set rsGetNewInvoice_cmd = Server.CreateObject ("ADODB.Command")
rsGetNewInvoice_cmd.ActiveConnection = MM_bodyartforms_sql_STRING
rsGetNewInvoice_cmd.CommandText = "SELECT ID, email FROM dbo.sent_items WHERE email = '" & (GetNotes.Fields.Item("email").Value) & "' ORDER BY ID DESC" 
rsGetNewInvoice_cmd.Prepared = true

Set rsGetNewInvoice = rsGetNewInvoice_cmd.Execute
rsGetNewInvoice_numRows = 0
%>
<% if request.form("ReturnMailer") = "RETURN ENVELOPE" then

set rsAddOrderDetail = Server.CreateObject("ADODB.Recordset")
rsAddOrderDetail.ActiveConnection = MM_bodyartforms_sql_STRING
rsAddOrderDetail.Source = "SELECT * FROM TBL_OrderSummary"
rsAddOrderDetail.CursorLocation = 3 'adUseClient
rsAddOrderDetail.LockType = 1 'Read-only records
rsAddOrderDetail.Open()
rsAddOrderDetail_numRows = 0

rsAddOrderDetail.addnew
rsAddOrderDetail("InvoiceID") = rsGetNewInvoice.Fields.Item("ID").Value
rsAddOrderDetail("ProductID") = 2991
rsAddOrderDetail("DetailID") = 17999
rsAddOrderDetail("qty") = 1
rsAddOrderDetail("item_price") = 0
rsAddOrderDetail.update

end if %>
<% 
set CopyItem = Server.CreateObject("ADODB.Command")
CopyItem.ActiveConnection = MM_bodyartforms_sql_STRING
CopyItem.CommandText = "INSERT INTO TBL_OrderSummary (InvoiceID, ProductID, DetailID, qty) SELECT " & (rsGetNewInvoice.Fields.Item("ID").Value) & ", ProductID, DetailID, qty FROM TBL_OrderSummary WHERE OrderDetailID =" & Request.Form("OrderDetailID") 
CopyItem.Execute() 
%>
<%
set commUpdate2 = Server.CreateObject("ADODB.Command")
commUpdate2.ActiveConnection = MM_bodyartforms_sql_STRING
CommUpdate2.CommandText = "UPDATE TBL_OrderSummary SET notes = 'Ref # " & (rsGetNewInvoice.Fields.Item("ID").Value) & "' WHERE OrderDetailID = " & Request.Form("OrderDetailID") & "" 
commUpdate2.Execute()
%>
</body>
</html>
<%
response.redirect "invoice.asp?ID=" & (rsGetNewInvoice.Fields.Item("ID").Value)
GetNotes.Close()
Set GetNotes = Nothing
%>
<%
rsGetNewInvoice.Close()
Set rsGetNewInvoice = Nothing
%>
