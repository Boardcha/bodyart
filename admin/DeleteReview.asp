<%@LANGUAGE="VBSCRIPT"%> 
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
set Command1 = Server.CreateObject("ADODB.Command")'create command object
Command1.ActiveConnection = MM_bodyartforms_sql_STRING 'connection string
Command1.CommandText = "DELETE FROM TBLReviews WHERE ReviewID = " & Request.QueryString("Review")
Command1.Execute() 

Response.Redirect("../productdetails.asp?ProductID=" & Request.QueryString("ID"))
%>

<html>
<head>
<title>Deleting product</title>
</head>
<body bgcolor="#666699" text="#CCCCCC">
Items are being deleted... 
</body>
</html>