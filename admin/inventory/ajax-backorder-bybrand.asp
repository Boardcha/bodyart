<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" %>
<!--#include virtual="/Connections/bodyartforms_sql_ADMIN.asp" -->
<%
var_brand = request.querystring("brand")

Set objCmd = Server.CreateObject ("ADODB.Command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT JEW.title, DET.Gauge, JEW.ProductID, JEW.picture, ORS.OrderDetailID, DET.length, SNT.ID, count(*) as qty " & _
						"FROM jewelry JEW " & _
						"INNER JOIN ProductDetails DET ON JEW.ProductID = DET.ProductID " & _
						"INNER JOIN TBL_OrderSummary AS ORS ON DET.ProductDetailID = ORS.DetailID " & _
						"INNER JOIN sent_items SNT ON SNT.ID = ORS.InvoiceID " & _
						"WHERE  (ORS.backorder = 1) AND brandname= ? AND JEW.active = 1 AND DET.active = 1" & _
						"GROUP BY title,  JEW.ProductID, JEW.picture, DET.Gauge, ORS.OrderDetailID, DET.length, SNT.ID " & _
						"ORDER BY JEW.ProductID, SNT.ID"
objCmd.Parameters.Append(objCmd.CreateParameter("value",200,1,50, var_brand))
Set rsGetBackorderList = objCmd.Execute()
%>
<div class="admin-content">
<% if not rsGetBackorderList.eof then %>
<table class="admin-table">
<thead>
<tr>
	<th colspan="2"><h1><%= var_brand %> - Backorder List</h1></th>
</tr>
<tr>
	<th>Invoice</th>
	<th>Item</th>
	<th>Gauge</th>
	<th>Length</th>
	<th class="pl-5 pr-5">Qty</th>
</tr>
</thead>
<% while not rsGetBackorderList.eof %>
<tr>
<td>
<a href="/admin/invoice.asp?id=<%= rsGetBackorderList("ID") %>" target="_blank"><%= rsGetBackorderList("ID") %></a>
</td>
<td>
<a href="/productdetails.asp?ProductID=<%= rsGetBackorderList("ProductID") %>" target="_blank"><img src="http://bodyartforms-products.bodyartforms.com/<%= rsGetBackorderList("picture") %>" class="mini-thumbnail" align="middle"></a>
&nbsp;&nbsp;
<%= rsGetBackorderList("title") %></td>
<td><%= rsGetBackorderList("gauge") %></td>
<td><%= rsGetBackorderList("length") %></td>
<td class="text-center">
<%= rsGetBackorderList("qty") %>
</td>

</tr>
<%
rsGetBackorderList.movenext()
wend
%>
</table>
<% end if %>
</div>
<%
DataConn.Close()
%>