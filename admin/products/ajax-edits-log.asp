<%@LANGUAGE="VBSCRIPT" %>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT top(100) e.description, e.edit_date, u.name, j.title, e.invoice_id, d.Gauge, d.Length, d.ProductDetail1, j.picture, j.ProductID, e.detail_id, e.invoice_detail_id, e.product_id, e.user_id FROM TBL_AdminUsers AS u INNER JOIN tbl_edits_log AS e ON u.ID = e.user_id LEFT OUTER JOIN TBL_OrderSummary AS s ON e.invoice_detail_id = s.OrderDetailID LEFT OUTER JOIN sent_items AS i ON e.invoice_id = i.ID LEFT OUTER JOIN jewelry AS j ON e.product_id = j.ProductID LEFT OUTER JOIN ProductDetails AS d ON e.detail_id = d.ProductDetailID WHERE (e.edit_date <> '') AND e.detail_id = ? AND invoice_detail_id = 0 ORDER BY e.edit_date DESC"
	objCmd.Parameters.Append(objCmd.CreateParameter("detailid",3,1,10,request("detailid")))
    set rs_getEdits = objCmd.Execute() 
%>

<% if NOT rs_getEdits.EOF then ' only show if recordset has results %>

<% 
if rs_getEdits.Fields.Item("ProductID").Value <> 0 then %>
<a href="product-edit.asp?ProductID=<%= rs_getEdits.Fields.Item("ProductID").Value %>" target="_blank"><img src="http://bodyartforms-products.bodyartforms.com/<%=(rs_getEdits.Fields.Item("picture").Value)%>" width="40px" style="padding-right: 10px;" align="left"></a><%= rs_getEdits.Fields.Item("title").Value %>
<% end if %>
<% 
if rs_getEdits.Fields.Item("detail_id").Value <> 0 then %>

&nbsp;&nbsp;&nbsp;|&nbsp;&nbsp; Item: <%= rs_getEdits.Fields.Item("Gauge").Value %>&nbsp;<%= rs_getEdits.Fields.Item("Length").Value %>&nbsp;<%= rs_getEdits.Fields.Item("ProductDetail1").Value %>

<% end if ' if detail id <> 0 %>
<table class="table table-striped table-borderless table-hover">
	<thead class="thead-dark">
		<tr>
			<th width="10%">Edited by</th>
			<th width="70%">Description</th>
			<th width="20%">Date edited</th>
		</tr>
	</thead>
<%
While NOT rs_getEdits.EOF
%>
	<tr>

		<td>
			<%= rs_getEdits.Fields.Item("name").Value %>
		</td>
		<td>
			<%= rs_getEdits.Fields.Item("description").Value %>
		</td>
		<td>
			<%= rs_getEdits.Fields.Item("edit_date").Value %>	
		</td>
	</tr>
<%
	rs_getEdits.MoveNext()
Wend
%>
</table>
<% else ' if no records found %>
No edits found
<% end if ' if recordset has results to show %>
</div><!-- end content area div -->
</body>
</html>
<%
Set rs_getEdits = Nothing
DataConn.Close()
%>
