<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
'==== PAGE HAS BEEN BOOSTRAPPED =======
bootstrapped = "yes"


Dim rsGetPreorders

Set rsGetPreorders = Server.CreateObject("ADODB.Recordset")
rsGetPreorders.ActiveConnection = MM_bodyartforms_sql_STRING
rsGetPreorders.Source = "SELECT TOP (100) PERCENT dbo.sent_items.date_sent, dbo.sent_items.email, dbo.sent_items.shipped, dbo.TBL_OrderSummary.item_received, dbo.TBL_OrderSummary.item_ordered, dbo.jewelry.brandname FROM dbo.jewelry INNER JOIN dbo.TBL_OrderSummary ON dbo.jewelry.ProductID = dbo.TBL_OrderSummary.ProductID INNER JOIN dbo.sent_items ON dbo.TBL_OrderSummary.InvoiceID = dbo.sent_items.ID WHERE (dbo.TBL_OrderSummary.item_ordered = 1) AND (dbo.sent_items.date_sent > '" & now() - 210 & "') AND (dbo.TBL_OrderSummary.item_received = 0) AND (dbo.jewelry.brandname = '" & request.querystring("brand") & "') AND dbo.sent_items.shipped = 'ON ORDER' ORDER BY dbo.sent_items.date_sent"
rsGetPreorders.CursorLocation = 3 'adUseClient
rsGetPreorders.LockType = 1 'Read-only records
rsGetPreorders.Open()

Dim PreviousDate, NewDate
Dim PreviousEmail, NewEmail
PreviousDate = ""
NewDate = ""
PreviousEmail = ""
NewEmail = ""
%>
<html>
<head>

<title>Pre-order e-mail lists</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1"></head>
<body>
<!--#include file="admin_header.asp"-->
<div class="p-3">
<h5>Customer e-mails for items not received (<%= rsGetPreorders.RecordCount %>)</h5>

<a href="preorder_emails.asp?brand=Anatometal">Anatometal</a> |  
<a href="preorder_emails.asp?brand=Industrial Strength">Industrial Strength</a> |  
<a href="preorder_emails.asp?brand=Neometal">Neometal</a><br>

<%
If rsGetPreorders.EOF Then
%>
<div class="alert alert-danger">No items found</div>
<%
	Else
  	Do While Not rsGetPreorders.EOF
	
	NewDate = rsGetPreorders.Fields.Item("date_sent").Value
	NewEmail = rsGetPreorders.Fields.Item("email").Value

    If NewDate <> PreviousDate Then
%>
<br>
<br>
        <span class="topnavlinks">Placed on <%= NewDate %></span><br>

<%
    End If
  
		If NewEmail <> PreviousEmail Then  
  %>
  
<%=(rsGetPreorders.Fields.Item("email").Value)%>;

<%
		End if
	PreviousDate = NewDate
	PreviousEmail = NewEmail
	rsGetPreorders.MoveNext()
	Loop
End If
%>
</div>
</body>
</html>
<%
rsGetPreorders.Close()
Set rsGetPreorders = Nothing
%>
