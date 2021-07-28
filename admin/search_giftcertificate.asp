<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
'==== PAGE HAS BEEN BOOSTRAPPED =======
bootstrapped = "yes"


if request.form("GiftCert") <> "" then
	var_cert_code = request.form("GiftCert")
	var_sql_code = "WHERE code = ?"
else
	var_cert_code = ""
	var_sql_code = ""
	
	if request.form("date_begin") <> "" and request.form("date_end") <> "" then
		var_sql_code = "WHERE date_created >= ? and date_created <= ?"
		var_date_display = "Showing certificates between " & FormatDateTime(request.form("date_begin"),2) & " and " & FormatDateTime(request.form("date_end"),2)
	else
		var_date_display = "Showing newest 25 gift certificates"
	end if
end if

Set objCmd = Server.CreateObject ("ADODB.Command")
objCmd.ActiveConnection = MM_bodyartforms_sql_STRING
objCmd.CommandText = "SELECT TOP(25) * FROM dbo.TBLcredits " & var_sql_code & " ORDER BY ID DESC" 
	if request.form("GiftCert") <> "" then
		objCmd.Parameters.Append objCmd.CreateParameter("code", 200, 1, 50, var_cert_code)
	end if
	if request.form("date_begin") <> "" and request.form("date_end") <> "" then
		objCmd.Parameters.Append(objCmd.CreateParameter("date_begin",200,1,50,request.form("date_begin")))
		objCmd.Parameters.Append(objCmd.CreateParameter("date_end",200,1,50,request.form("date_end")))
	end if
Set rsGetCertificate = objCmd.Execute()
%>
<html>
<head>
<title>Gift certificate history</title>
</head>
<body>
<!--#include file="admin_header.asp"-->
<div class="p-3">
<h5><%= var_date_display %></h5>
<div class="my-2">
	Search by date range:
</div>
<form class="form-inline">
<label class="mr-2">Start</label><input class="form-control form-control-sm mr-4" type="date" name="date_begin" id="date_begin">
<label class="mr-2">End</label><input class="form-control form-control-sm" type="date" name="date_end">

<button class="btn btn-sm btn-secondary ml-3" type="submit" formaction="search_giftcertificate.asp" formmethod="post">Search</button>
</form>

<% If Not rsGetCertificate.EOF Then 
While NOT rsGetCertificate.EOF 

Set objCmd = Server.CreateObject ("ADODB.Command")
objCmd.ActiveConnection = MM_bodyartforms_sql_STRING
objCmd.CommandText = "SELECT * FROM dbo.TBL_Credits_UsedOn WHERE OriginalCreditID = " & rsGetCertificate.Fields.Item("ID").Value & "" 
Set rsGetOrdersUsedOn = objCmd.Execute()
%>
<hr>
<h5 class="text-info">
	<%= rsGetCertificate.Fields.Item("code").Value %>
	<span class="ml-5">Amount left <%= formatcurrency(rsGetCertificate.Fields.Item("amount").Value, -1, -2, -0, -2) %></span>
	<a class="btn btn-sm btn-secondary ml-4" href="giftcertificate.asp?id=<%= rsGetCertificate.Fields.Item("invoice").Value %>">Edit</a>
</h5>
<%
if rsGetCertificate.Fields.Item("custid_converted").Value <> 0 then
%> 
Converted to store credit by customer ID <a href="customer_edit.asp?ID=<%= rsGetCertificate.Fields.Item("custid_converted").Value %>" target="_blank"><%= rsGetCertificate.Fields.Item("custid_converted").Value %></a><br/>
<% end if %>
Created on <%= rsGetCertificate.Fields.Item("date_created").Value %>
<br/>
Purchased on invoice <a href="invoice.asp?ID=<%= rsGetCertificate.Fields.Item("invoice").Value %>" target="_blank"><%= rsGetCertificate.Fields.Item("invoice").Value %></a>
<br/>
Recipients name: <%=(rsGetCertificate.Fields.Item("rec_name").Value)%><br/>
Recipients email: <%=(rsGetCertificate.Fields.Item("rec_email").Value)%>

<% if NOT rsGetOrdersUsedOn.EOF then %>
<div class="mt-3">
Used on orders:
<% 
While NOT rsGetOrdersUsedOn.EOF
%>
        <a href="invoice.asp?ID=<%=(rsGetOrdersUsedOn.Fields.Item("InvoiceUsedOn").Value)%>" target="_blank"><%=(rsGetOrdersUsedOn.Fields.Item("InvoiceUsedOn").Value)%></a>
        <% 
  rsGetOrdersUsedOn.MoveNext()
Wend
%>
</div>
<%
	end if ' if used on orders are found
  rsGetCertificate.MoveNext()
Wend

else ' if rsGetCertificate.EOF  
%>
	No gift certificate found
<% End If ' if rsGetCertificate.EOF 
rsGetCertificate.Close() 
%>
</div>
</body>
</html>