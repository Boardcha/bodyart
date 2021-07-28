<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
Dim rsSetID__MMColParam
rsSetID__MMColParam = "1"
If (Request.Form("email") <> "") Then 
  rsSetID__MMColParam = Request.Form("email")
End If
%>
<%
Dim rsSetID
Dim rsSetID_numRows

Set rsSetID = Server.CreateObject("ADODB.Recordset")
rsSetID.ActiveConnection = MM_bodyartforms_sql_STRING
rsSetID.Source = "SELECT customer_ID, ID, customer_first, customer_last, email FROM dbo.sent_items WHERE email = '" + rsSetID__MMColParam + "'"
rsSetID.CursorLocation = 3 'adUseClient
rsSetID.LockType = 1 'Read-only records
rsSetID.Open()

rsSetID_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = -1
Repeat1__index = 0
rsSetID_numRows = rsSetID_numRows + Repeat1__numRows
%>

<html>
<head>
<link rel="stylesheet" type="text/css" href="../includes/nav.css" />
<title>Edit customer info</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body bgcolor="#666699" text="#CCCCCC" link="#CCCCCC" vlink="#CCCCCC" topmargin="0" class="pricegauge">
<!--#include file="admin_header.asp"-->
<span class="adminheader">Searching customer orders</span><P>
<% If Not rsSetID.EOF Or Not rsSetID.BOF Then %>
        <% 
While ((Repeat1__numRows <> 0) AND (NOT rsSetID.EOF)) 
%>
            Invoice # <%=(rsSetID.Fields.Item("ID").Value)%> ... updated 
            <%
set commUpdate = Server.CreateObject("ADODB.Command")
commUpdate.ActiveConnection = MM_bodyartforms_sql_STRING
commUpdate.CommandText = "UPDATE sent_items SET customer_ID = "&Request.form("custID")&" WHERE ID = "&rsSetID.Fields.Item("ID").Value
commUpdate.Execute()

%><br>

            <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rsSetID.MoveNext()
Wend
%>
        <p>All past orders with email <%= Request.form("email") %> AND first name <%= Request.Form("first") %> have been added to your account.</p>
        <% End If ' end Not rsSetID.EOF Or NOT rsSetID.BOF %>
      <% If rsSetID.EOF And rsSetID.BOF Then %>
        <p>Sorry but no orders with email <%= Request.form("email") %> AND first name <%= Request.Form("first") %> were found in the system to add to your account.</p>
        <% End If ' end rsSetID.EOF And rsSetID.BOF %>
<p>&nbsp;</p>
</body>
</html>
<%
rsSetID.Close()
Set rsSetID = Nothing
%>
