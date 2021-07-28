<%@LANGUAGE="VBSCRIPT"%> 
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
set commUpdate = Server.CreateObject("ADODB.Command")
commUpdate.ActiveConnection = MM_bodyartforms_sql_STRING

commUpdate.CommandText = "UPDATE TBL_OrderSummary SET ErrorOnReview = 0, ErrorDescription = item_problem + ' Dismissed by " & rsGetUser.Fields.Item("name").Value & " " & now() &"', item_problem = '0' WHERE OrderDetailID = " & Request.Querystring("DetailID") 
    ' comment out next line AFTER IT WORKS 
    'Response.Write "DEBUG SQL: " & commUpdate.CommandText & "<BR/>" 
commUpdate.Execute()
Set commUpdate = Nothing

   Response.Redirect("invoice.asp?ID=" & Request.Querystring("ID") & "")
%>
<html>
<head>
<title>Records are being updated</title>
</head>
<body bgcolor="#666699" text="#CCCCCC">
Records are being updated
</body>
</html>
