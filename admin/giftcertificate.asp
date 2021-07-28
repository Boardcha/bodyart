<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
Dim getCert__MMColParam
getCert__MMColParam = "1"
If (Request.QueryString("ID") <> "") Then 
  getCert__MMColParam = Request.QueryString("ID")
End If
%>
<%
Dim getCert
Dim getCert_numRows

Set getCert = Server.CreateObject("ADODB.Recordset")
getCert.ActiveConnection = MM_bodyartforms_sql_STRING
getCert.Source = "SELECT * FROM dbo.TBLcredits WHERE invoice = '" + Replace(getCert__MMColParam, "'", "''") + "'"
' Original CursorLocation
getCert.CursorType = 1
getCert.CursorLocation = 2
getCert.LockType = 3
getCert.Open()

getCert_numRows = 0
%>
<%
Dim rsGetInvoice__MMColParam
rsGetInvoice__MMColParam = "1"
If (Request.QueryString("ID") <> "") Then 
  rsGetInvoice__MMColParam = Request.QueryString("ID")
End If
%>
<%
Dim rsGetInvoice
Dim rsGetInvoice_numRows

Set rsGetInvoice = Server.CreateObject("ADODB.Recordset")
rsGetInvoice.ActiveConnection = MM_bodyartforms_sql_STRING
rsGetInvoice.Source = "SELECT ID, customer_first, customer_last, email FROM dbo.sent_items WHERE ID = " + Replace(rsGetInvoice__MMColParam, "'", "''") + ""
rsGetInvoice.CursorLocation = 3 'adUseClient
rsGetInvoice.LockType = 1 'Read-only records
rsGetInvoice.Open()

rsGetInvoice_numRows = 0
%>
<%
Dim rsGetOrder__MMColParam
rsGetOrder__MMColParam = "1"
If (Request.QueryString("ID") <> "") Then 
  rsGetOrder__MMColParam = Request.QueryString("ID")
End If
%>
<%
Dim rsGetOrder
Dim rsGetOrder_numRows

Set rsGetOrder = Server.CreateObject("ADODB.Recordset")
rsGetOrder.ActiveConnection = MM_bodyartforms_sql_STRING
rsGetOrder.Source = "SELECT ID, customer_first, customer_last, email FROM dbo.sent_items WHERE ID = " + Replace(rsGetOrder__MMColParam, "'", "''") + ""
rsGetOrder.CursorLocation = 3 'adUseClient
rsGetOrder.LockType = 1 'Read-only records
rsGetOrder.Open()

rsGetOrder_numRows = 0
%>
<html>
<head>
<title>Gift certificate email</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body>
  <!--#include file="admin_header.asp"-->

<form class="p-3" name="form1" method="post" action="">
  <% if request.form("send") = "yes" then %>
<div class="alert alert-success">EMAIL HAS BEEN SENT</div>
<% END IF %>
  

    <span class="pricegauge">Senders name:</span>
   
      <input name="rec_yourname" type="text" class="adminfields" id="rec_yourname" value="<%=(rsGetOrder.Fields.Item("customer_first").Value)%>&nbsp;<%=(rsGetOrder.Fields.Item("customer_last").Value)%>" size="35">
      
<br>
    <span class="pricegauge">Senders email:</span>
    <input name="your_email" type="text" class="adminfields" id="your_email" value="<%=(rsGetOrder.Fields.Item("email").Value)%>" size="35">
<br>
    <span class="pricegauge">Recipients name:</span>
    <input name="rec_name" type="text" class="adminfields" id="rec_name" value="<% If Not getCert.EOF Or Not getCert.BOF Then %><%=(getCert.Fields.Item("rec_name").Value)%><% End If ' end Not getCert.EOF Or NOT getCert.BOF %>">
    <br>
    <span class="pricegauge">Recipients email:</span>
    <input name="rec_email" type="text" class="adminfields" id="rec_email" value="<% If Not getCert.EOF Or Not getCert.BOF Then %><%=(getCert.Fields.Item("rec_email").Value)%><% End If ' end Not getCert.EOF Or NOT getCert.BOF %>" size="35">
    <br>
    <span class="pricegauge">Code:</span>
    <input name="code" type="text" class="adminfields" id="code" value="<% If Not getCert.EOF Or Not getCert.BOF Then %><%=(getCert.Fields.Item("code").Value)%><% End If ' end Not getCert.EOF Or NOT getCert.BOF %>">
     <span class="pricegauge">(6 digits mixed with numbers &amp; letters) </span><br>
    <span class="pricegauge">Amount:</span>
    <input name="gift_amount" type="text" class="adminfields" id="gift_amount" value="<% If Not getCert.EOF Or Not getCert.BOF Then %><%=(getCert.Fields.Item("amount").Value)%><% End If ' end Not getCert.EOF Or NOT getCert.BOF %>">
    <br>
    <span class="pricegauge">Message:</span>
    <textarea name="rec_message" cols="35" class="adminfields" id="rec_message"><% If Not getCert.EOF Or Not getCert.BOF Then %><%=(getCert.Fields.Item("message").Value)%><% End If ' end Not getCert.EOF Or NOT getCert.BOF %></textarea>
    <input name="send" type="hidden" id="send" value="yes">
    <% If Not getCert.EOF Or Not getCert.BOF Then %>
      <input name="Update" type="hidden" id="Update" value="yes">
      <input name="CertID" type="hidden" id="CertID" value="<%=(getCert.Fields.Item("ID").Value)%>">
      <% End If ' end Not getCert.EOF Or NOT getCert.BOF %>
<% If getCert.EOF And getCert.BOF Then %>
      <input name="AddNew" type="hidden" id="AddNew" value="yes">
      <% End If ' end getCert.EOF And getCert.BOF %>
</p>
<input class="btn btn-purple" type="submit" name="Submit" value="Submit">
</form>
<% if request.form("AddNew") = "yes" then 

set rsAddRecord = Server.CreateObject("ADODB.Recordset")
rsAddRecord.ActiveConnection = MM_bodyartforms_sql_STRING
rsAddRecord.Source = "SELECT * FROM TBLcredits"
' Original CursorLocation
rsAddRecord.CursorType = 1
rsAddRecord.CursorLocation = 2
rsAddRecord.LockType = 3
rsAddRecord.Open()
rsAddRecord_numRows = 0

rsAddRecord.addnew
rsAddRecord("invoice") = request.querystring("ID")
rsAddRecord("name") = request.form("rec_yourname")
rsAddRecord("rec_name") = request.form("rec_name")
rsAddRecord("rec_email") = request.form("rec_email")
rsAddRecord("code") = request.form("code")
rsAddRecord("amount") = request.form("gift_amount")
rsAddRecord("message") = request.form("rec_message")
rsAddRecord.update

end if %>

<% if request.form("Update") = "yes" then 

set commUpdate = Server.CreateObject("ADODB.Command")
commUpdate.ActiveConnection = MM_bodyartforms_sql_STRING
commUpdate.CommandText = "UPDATE TBLcredits SET invoice = '" + request.querystring("ID") + "', name = '" + request.form("rec_yourname") + "', rec_name = '" + request.form("rec_name") + "', rec_email = '" + request.form("rec_email") + "', code = '" + request.form("code") + "', amount = '" + request.form("gift_amount") + "', message = '" + request.form("rec_message") + "' WHERE ID = " & request.form("CertID") & "" 
commUpdate.Execute()

 end if %>

<% if request.form("send") = "yes" then

done_mailing_certs = "no"

your_name = request.form("rec_yourname")
gift_amount = request.form("gift_amount")
rec_email = request.form("rec_email")
rec_name = request.form("rec_name")
var_cert_code = request.form("code")
message = request.form("rec_message")
%>
<!--#include virtual="/emails/function-send-email.asp"-->
<!--#include virtual="/emails/email_variables.asp"-->
<%

end if %>
</body>
</html>
<%
getCert.Close()
Set getCert = Nothing
%>
<%
rsGetInvoice.Close()
Set rsGetInvoice = Nothing
%>
<%
rsGetOrder.Close()
Set rsGetOrder = Nothing
%>
