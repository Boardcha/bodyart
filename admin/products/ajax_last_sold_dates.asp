<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" %>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
if request.form("detailid") <> "" then
'response.write request.form("detailid")

	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT TOP (20) sent_items.ID, date_order_placed, qty FROM TBL_OrderSummary INNER JOIN sent_items ON TBL_OrderSummary.InvoiceID = sent_items.ID WHERE (sent_items.ship_code = N'paid') AND (TBL_OrderSummary.DetailID = ?) ORDER BY sent_items.date_order_placed DESC"
	objCmd.Parameters.Append(objCmd.CreateParameter("detailid",3,1,10,request.form("detailid")))
	Set rsGetDatesSold = objCmd.Execute()

end if
%>
<table class="table table-sm">
	<thead class="thead-dark">
	  <tr>
		<th class="py-0" scope="col">Date sold</th>
		<th class="py-0 text-center" scope="col">Qty sold</th>
	  </tr>
	</thead>
	<tbody>
<%
While NOT rsGetDatesSold.EOF 
%>
	  <tr>
		<td class="py-0">
			<a class="mr-3" href='invoice.asp?ID=<%=(rsGetDatesSold.Fields.Item("ID").Value)%>' target='_blank' ><%=FormatDateTime((rsGetDatesSold.Fields.Item("date_order_placed").Value), 2)%></a>
		</td>
		<td class="py-0 text-center">
			<%= rsGetDatesSold.Fields.Item("qty").Value %>
		</td>
	  </tr>
  <% 
  rsGetDatesSold.MoveNext()
Wend
%>
</tbody>
</table>

<%
set rsGetDatesSold = nothing
DataConn.Close()
%>