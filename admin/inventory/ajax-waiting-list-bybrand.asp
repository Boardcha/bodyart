<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" %>
<!--#include virtual="/Connections/bodyartforms_sql_ADMIN.asp" -->
<%
var_brand = request.querystring("brand")

Set objCmd = Server.CreateObject ("ADODB.Command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT TOP (100) PERCENT title, name, email, ProductID, picture FROM QRYWaitingList WHERE (brandname = ?) ORDER BY title, name" 
objCmd.Parameters.Append(objCmd.CreateParameter("value",200,1,50, var_brand))
Set rsGetWaitingList = objCmd.Execute()
%>
<div class="admin-content">
<% if not rsGetWaitingList.eof then %>
<table class="admin-table">
<thead>
<tr>
	<th colspan="3"><h1><%= var_brand %> waiting list</h1></th>
</tr>
<tr>
	<th>Item</th>
	<th>Customer name</th>
	<th>Customer e-mail</th>
</tr>
</thead>
<% while not rsGetWaitingList.eof %>
<tr>
<td>
<a href="/productdetails.asp?ProductID=<%= rsGetWaitingList.Fields.Item("ProductID").Value %>" target="_blank"><img src="http://bodyartforms-products.bodyartforms.com/<%= rsGetWaitingList.Fields.Item("picture").Value %>" class="mini-thumbnail" align="middle"></a>
&nbsp;&nbsp;
<%= rsGetWaitingList.Fields.Item("title").Value %></td>
<td>
<%= rsGetWaitingList.Fields.Item("name").Value %>
</td>
<td>
<%= rsGetWaitingList.Fields.Item("email").Value %>
</td>
</tr>
<%
rsGetWaitingList.movenext()
wend
%>
</table>
<% end if %>
</div>
<%
DataConn.Close()
%>