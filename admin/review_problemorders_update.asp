<%@LANGUAGE="VBSCRIPT"%> 
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
set commUpdate = Server.CreateObject("ADODB.Command")
commUpdate.ActiveConnection = MM_bodyartforms_sql_STRING

commUpdate.CommandText = "UPDATE sent_items SET Review_OrderError = 0, shipped = '" & Request.form("status") & "', our_notes = '" + Request.form("our_notes") + "', AmountLost = " & Request.form("AmountLost") & " WHERE ID = " & Request.form("InvoiceID") 
    ' comment out next line AFTER IT WORKS 
    'Response.Write "DEBUG SQL: " & commUpdate.CommandText & "<BR/>" 
commUpdate.Execute()
Set commUpdate = Nothing

set commUpdate = Server.CreateObject("ADODB.Command")
commUpdate.ActiveConnection = MM_bodyartforms_sql_STRING

commUpdate.CommandText = "UPDATE TBL_OrderSummary SET notes = '" & Request.form("notes") & "' WHERE OrderDetailID = " & Request.Form("OrderDetailID") 
    ' comment out next line AFTER IT WORKS 
    'Response.Write "DEBUG SQL: " & commUpdate.CommandText & "<BR/>" 
commUpdate.Execute()
Set commUpdate = Nothing

   Response.Redirect("review_problemorders.asp")
%>
<html>
<head>
<title>Records are being updated</title>
</head>
<body bgcolor="#666699" text="#CCCCCC">
Records are being updated
</body>
</html>
