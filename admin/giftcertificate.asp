<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
if request.form("Update") = "yes" then 

  set objCmd = Server.CreateObject("ADODB.Command")
  objCmd.ActiveConnection = MM_bodyartforms_sql_STRING
  objCmd.CommandText = "UPDATE TBLcredits SET invoice = ?, rec_name = ?, rec_email = ?, code = ?, amount = ?, message = ? WHERE ID = ?" 
  objCmd.Parameters.Append(objCmd.CreateParameter("invoice",3,1,15, request.querystring("ID") ))
  objCmd.Parameters.Append(objCmd.CreateParameter("rec_name",200,1,25, request.form("rec_name") ))
  objCmd.Parameters.Append(objCmd.CreateParameter("rec_email",200,1,50, request.form("rec_email") ))
  objCmd.Parameters.Append(objCmd.CreateParameter("code",200,1,50, request.form("code") ))
  objCmd.Parameters.Append(objCmd.CreateParameter("amount",6,1,10, request.form("gift_amount") ))
  objCmd.Parameters.Append(objCmd.CreateParameter("rec_message",200,1,250, request.form("rec_message") ))
  objCmd.Parameters.Append(objCmd.CreateParameter("certificate_id",3,1,15, request.form("CertID") ))
  objCmd.Execute()

end if 


Set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = MM_bodyartforms_sql_STRING
objCmd.CommandText = "SELECT * FROM dbo.TBLcredits WHERE invoice = ?"
objCmd.Parameters.Append(objCmd.CreateParameter("invoice",3,1,15, request.querystring("ID") ))
Set getCert = objCmd.Execute()


Set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = MM_bodyartforms_sql_STRING
objCmd.CommandText = "SELECT ID, customer_first, customer_last, email FROM dbo.sent_items WHERE ID = ?"
objCmd.Parameters.Append(objCmd.CreateParameter("invoice",3,1,15, request.querystring("ID") ))
Set rsGetOrder = objCmd.Execute()
%>
<html>
<head>
<title>Gift certificate email</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body>
  <!--#include file="admin_header.asp"-->
<style>
  label{margin:0}
  .form-control-sm {margin-bottom:15px}
</style>
<form class="p-3" name="form1" method="post" action="">
  <% if request.form("send") = "yes" then %>
<div class="alert alert-success">EMAIL HAS BEEN SENT</div>
<% END IF %>
  

<div class="font-weight-bold">Sending from:</div>
<%=(rsGetOrder.Fields.Item("customer_first").Value)%>&nbsp;<%=(rsGetOrder.Fields.Item("customer_last").Value)%>  |  
<%=(rsGetOrder.Fields.Item("email").Value)%>
<br><br>
  <input name="rec_yourname" type="hidden" id="rec_yourname" value="<%= rsGetOrder("customer_first")  & " " & rsGetOrder("customer_last") %>">
  <input name="your_email" type="hidden" id="your_email" value="<%=(rsGetOrder.Fields.Item("email").Value)%>">

    <label>Recipients name:</label>
    <input name="rec_name" type="text" class="form-control form-control-sm" id="rec_name" value="<% If Not getCert.EOF Or Not getCert.BOF Then %><%=(getCert.Fields.Item("rec_name").Value)%><% End If ' end Not getCert.EOF Or NOT getCert.BOF %>">

    <label>Recipients email:</label>
    <input name="rec_email" type="text" class="form-control form-control-sm" id="rec_email" value="<% If Not getCert.EOF Or Not getCert.BOF Then %><%=(getCert.Fields.Item("rec_email").Value)%><% End If ' end Not getCert.EOF Or NOT getCert.BOF %>" size="35">

    <label>Code:</label>
    <input name="code" type="text" class="form-control form-control-sm" id="code" value="<% If Not getCert.EOF Or Not getCert.BOF Then %><%=(getCert.Fields.Item("code").Value)%><% End If ' end Not getCert.EOF Or NOT getCert.BOF %>">

    <label>Amount:</label>
    <input name="gift_amount" type="text" class="form-control form-control-sm" id="gift_amount" value="<% If Not getCert.EOF Or Not getCert.BOF Then %><%=(getCert.Fields.Item("amount").Value)%><% End If ' end Not getCert.EOF Or NOT getCert.BOF %>">

    <label>Message:</label>
    <textarea name="rec_message" cols="35" class="form-control form-control-sm" id="rec_message"><% If Not getCert.EOF Or Not getCert.BOF Then %><%=(getCert.Fields.Item("message").Value)%><% End If ' end Not getCert.EOF Or NOT getCert.BOF %></textarea>
    <input name="send" type="hidden" id="send" value="yes">
    <% If Not getCert.EOF Or Not getCert.BOF Then %>
      <input name="Update" type="hidden" id="Update" value="yes">
      <input name="CertID" type="hidden" id="CertID" value="<%=(getCert.Fields.Item("ID").Value)%>">
      <% End If ' end Not getCert.EOF Or NOT getCert.BOF %>
<% If getCert.EOF And getCert.BOF Then %>
      <input name="AddNew" type="hidden" id="AddNew" value="yes">
      <% End If ' end getCert.EOF And getCert.BOF %>

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

end if 



 if request.form("send") = "yes" then

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

rsGetOrder.Close()
Set rsGetOrder = Nothing
%>
