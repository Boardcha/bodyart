<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
'==== PAGE HAS BEEN BOOSTRAPPED =======
bootstrapped = "yes"
%>
<html>
<head>
<title>Add a new gift certificate</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body>
<!--#include file="admin_header.asp"-->
<div class="p-3">
<h5>
  Create a new gift certificate code
</h5>
<form name="form1" method="post" action="">

    <% if request.form("send") = "yes" then %>
<div class="alert alert-success">EMAIL HAS BEEN SENT</div>
<% END IF %>

<div class="form-group">
    <label for="rec_yourname">Senders name:</label>
      <input class="form-control form-control-sm" style="width:300px" name="rec_yourname" type="text"  id="rec_yourname" size="35">
</div>
  <div class="form-group">
    <label for="your_email">Senders email:</label>
    <input class="form-control form-control-sm" style="width:300px" name="your_email" type="text" id="your_email" size="35">
  </div>

  <div class="form-group">
    <label class="pricegauge">Recipients name:</label>
    <input class="form-control form-control-sm" style="width:300px" name="rec_name" type="text" id="rec_name">
  </div>

  <div class="form-group">
    <label class="pricegauge">Recipients email:</label>
    <input class="form-control form-control-sm" style="width:300px" name="rec_email" type="text" id="rec_email" size="35">
  </div>

  <div class="form-group">
    <label class="pricegauge">Code:</label>
    <input class="form-control form-control-sm" style="width:300px" name="code" type="text" id="code">
  </div>

  <div class="form-group">
    <label class="pricegauge">Amount:</label>
    <input class="form-control form-control-sm" style="width:300px" name="gift_amount" type="text" id="gift_amount">
  </div>

  <div class="form-group">
    <label class="pricegauge">Message:</label>
    <textarea class="form-control form-control-sm" style="width:300px" name="rec_message" cols="35" id="rec_message"></textarea>
  </div>
    
    <input name="send" type="hidden" id="send" value="yes">
    <input name="AddNew" type="hidden" id="AddNew" value="yes">
    <button class="btn btn-sm btn-secondary" type="submit" name="Submit">Create code</button>
</form>
<% if request.form("AddNew") = "yes" then 

set rsAddRecord = Server.CreateObject("ADODB.Recordset")
rsAddRecord.ActiveConnection = MM_bodyartforms_sql_STRING
rsAddRecord.Source = "SELECT * FROM TBLcredits"
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
rsAddRecord("certificate_original_amount") = request.form("gift_amount")
rsAddRecord("message") = request.form("rec_message")
rsAddRecord.update

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
</div>
</body>
</html>
