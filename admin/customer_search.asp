<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
'==== PAGE HAS BEEN BOOSTRAPPED =======
bootstrapped = "yes"


Dim rsGetCustomer
Dim rsGetCustomer_numRows

Set rsGetCustomer = Server.CreateObject("ADODB.Recordset")
rsGetCustomer.ActiveConnection = MM_bodyartforms_sql_STRING
rsGetCustomer.Source = "SELECT * FROM customers WHERE customer_first = '"+ Request.Querystring("first") +"' AND customer_last = '"+ Request.Querystring("last") +"' OR email = '"+ Request.Querystring("email") +"' OR customer_ID = '"+ Request.Querystring("CustomerID") +"'"
rsGetCustomer.CursorLocation = 3 'adUseClient
rsGetCustomer.LockType = 1 'Read-only records
rsGetCustomer.Open()

rsGetCustomer_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = -1
Repeat1__index = 0
rsGetCustomer_numRows = rsGetCustomer_numRows + Repeat1__numRows
%>
<%
if request.querystring("task") = "delete" then

set Command1 = Server.CreateObject("ADODB.Command")'create command object
Command1.ActiveConnection = MM_bodyartforms_sql_STRING 'connection string
Command1.CommandText = "DELETE FROM customers WHERE customer_ID = " & Request.Querystring("ID")
Command1.Execute()

Response.Write "<b>CUSTOMER HAS BEEN DELETED</b>"
Response.Redirect "customer_search.asp?first=" + Request.Querystring("first") + "&last=" + Request.Querystring("last") + "&email=" + Request.Querystring("email") + ""
End if
%>
<html>
<title>Customer search</title>
</head>
<body>
<!--#include file="admin_header.asp"-->
<div class="p-3">
<h4 class="mb-3">
	Customer accounts
</h4>
<% 
While NOT rsGetCustomer.EOF
%>
<hr>
<a class="btn btn-sm btn-danger mr-3" href="customer_search.asp?first=<%=Request.Querystring("first")%>&last=<%=Request.Querystring("last")%>&email=<%=Request.Querystring("email")%>&ID=<%=(rsGetCustomer.Fields.Item("customer_ID").Value)%>&task=delete">Delete</a>

<a class="btn btn-sm btn-secondary mr-3" href="customer_edit.asp?ID=<%=(rsGetCustomer.Fields.Item("customer_ID").Value)%>">Edit account</a><a href="customer_edit.asp?ID=<%=(rsGetCustomer.Fields.Item("customer_ID").Value)%>"></a>

<a class="btn btn-sm btn-secondary mr-3" href="order history.asp?var_first=<%=(rsGetCustomer.Fields.Item("customer_first").Value)%>&var_last=<%=(rsGetCustomer.Fields.Item("customer_last").Value)%>" >Orders by first/last</a>

<a class="btn btn-sm btn-secondary mr-3" href="order history.asp?custid=<%=(rsGetCustomer.Fields.Item("customer_ID").Value)%>">Orders by customer ID</a>

<div class="mt-2">
      Customer ID # <%=(rsGetCustomer.Fields.Item("customer_ID").Value)%><br>
          <%=(rsGetCustomer.Fields.Item("customer_first").Value)%> <%=(rsGetCustomer.Fields.Item("customer_last").Value)%><br/>
          <%=(rsGetCustomer.Fields.Item("email").Value)%><br>
In-store credit: $<%=(rsGetCustomer.Fields.Item("credits").Value)%> 
</div>
<% 
  rsGetCustomer.MoveNext()
Wend
%>

</div>
</body>
</html>
<%
rsGetCustomer.Close()
Set rsGetCustomer = Nothing
%>
