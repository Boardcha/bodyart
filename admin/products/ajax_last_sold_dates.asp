<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" %>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
if request.form("detailid") <> "" then
'response.write request.form("detailid")

	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT TOP (6) ORD.DetailID, Format(SNT.date_order_placed, 'yyyy-MM') as monthly, PDE.DateLastPurchased, PDE.qty AS onHand, SUM(ORD.qty) as qty_sold  FROM TBL_OrderSummary ORD INNER JOIN sent_items SNT ON ORD.InvoiceID = SNT.ID " & _
	"INNER JOIN ProductDetails PDE ON ORD.DetailID = PDE.ProductDetailID " & _
	"WHERE (SNT.ship_code = N'paid') AND (ORD.DetailID = ?) AND SNT.date_order_placed > DATEADD(MONTH, -6, GETDATE()) " & _
	"GROUP BY Format(SNT.date_order_placed, 'yyyy-MM'), ORD.DetailID, PDE.DateLastPurchased, PDE.qty " & _
	"ORDER BY monthly DESC"
	
	objCmd.Parameters.Append(objCmd.CreateParameter("detailid",3,1,10,request.form("detailid")))
	Set rsGetDatesSold = objCmd.Execute()

end if
%>
<table class="table table-sm">
	<thead class="thead-dark">
	  <tr>
		<th class="py-0" scope="col">Month</th>
		<th class="py-0 text-center" scope="col">Qty sold</th>
	  </tr>
	</thead>
	<tbody>
<%
If Not rsGetDatesSold.EOF Then
	While NOT rsGetDatesSold.EOF 
	%>
	<tr>
		<td class="py-0">
			<%=rsGetDatesSold.Fields.Item("monthly").Value%>
		</td>
		<td class="py-0 text-center">
			<%= rsGetDatesSold.Fields.Item("qty_sold").Value %>
		</td>
	</tr>
	  <% 
	  total = total + rsGetDatesSold.Fields.Item("qty_sold").Value 
	  rsGetDatesSold.MoveNext()
	Wend%>
	<tr>
		<td class="pt-3 py-0" colspan="2">
			(<%=total%>) sales in last 6 months.
		</td>
	</tr>	
<%Else%>
	<tr>
		<td class="pt-3 py-0" colspan="2">
			No sales in last 6 months.
		</td>
	</tr>
<%
End If	
%>
</tbody>
</table>

<%
set rsGetDatesSold = nothing
DataConn.Close()
%>