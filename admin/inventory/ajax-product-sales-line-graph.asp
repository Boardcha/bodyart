<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" %>
<!--#include virtual="/Connections/bodyartforms_sql_ADMIN.asp" -->
<%
detailID = request.querystring("detailID")
Set objCmd = Server.CreateObject ("ADODB.Command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText= "SELECT T3._month, COALESCE(SUM(total_qty), 0) As total_qty FROM (SELECT YEAR(date_added) * 100 + MONTH(date_added) As _month, qty As total_qty, date_added, DetailID, InvoiceID FROM TBL_OrderSummary) T1 " & _
					"INNER JOIN sent_items T2 ON T1.InvoiceID = T2.ID AND T2.ship_code = 'paid' " & _
					"RIGHT JOIN (SELECT DISTINCT YEAR(date_added) * 100 + MONTH(date_added) As _month FROM TBL_OrderSummary WHERE date_added > DATEADD(""m"", -12, GETDATE())) T3 " & _
					"ON T1._month = T3._month AND T1.date_added > DATEADD(""m"", -12, GETDATE()) AND T1.DetailID = " & detailID & _
					"GROUP BY T3._month " & _
					"ORDER BY T3._month"
Set rsSales = objCmd.Execute()

Set objCmd = Server.CreateObject ("ADODB.Command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText= "SELECT T1._month, COALESCE(T2.po_qty, 0) As po_qty FROM " & _
					"(SELECT DISTINCT YEAR(date_added) * 100 + MONTH(date_added) As _month FROM TBL_OrderSummary WHERE date_added > DATEADD(""m"", -12, GETDATE())) T1 " & _
					"LEFT JOIN  " & _
					"(SELECT COALESCE(SUM(po_qty), 0) As po_qty, YEAR(po_date_received) * 100 + MONTH(po_date_received) As _month FROM tbl_po_details  " & _
					"WHERE po_detailid = " & detailID & " AND po_date_received > DATEADD(""m"", -12, GETDATE()) AND po_received = 1 GROUP BY YEAR(po_date_received) * 100 + MONTH(po_date_received)) T2 ON T1._month = T2._month " & _
					"ORDER BY T1._month"
Set rsRestock = objCmd.Execute()

Set objCmd = Server.CreateObject ("ADODB.Command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT qty, DateLastPurchased, gauge, length, productdetail1 FROM productdetails WHERE ProductID = (Select TOP 1 ProductID FROM productdetails WHERE productdetailid= " & detailID & ")"
Set rsProductInfo = objCmd.Execute()

Dim month(12)
Dim sales(12)
Dim restock(12)
%>
<%	
If Not rsSales.EOF Then
		For i = 1 to 12
			month(i) = GetMonth(RIGHT(rsSales("_month"), 2))
			sales(i) = rsSales("total_qty")
			total_sales = total_sales + sales(i)
			rsSales.MoveNext 
		Next 
End If 
If Not rsRestock.EOF Then
		For i = 1 to 12
			restock(i) = rsRestock("po_qty")
			total_restock = total_restock + restock(i)
			rsRestock.MoveNext 
		Next 
End If
%>
<div class="col-md-6">
<canvas id="lineChart_<%=detailID%>"></canvas>
<%If Not rsProductInfo.EOF Then%>
	<table>
		<th>Qty</th>
		<th>Gauge</th>
		<th>Length</th>
		<th>DateLastPurchased</th>
		<th>Product Detail</th>
		<%While Not rsProductInfo.EOF%>
			<tr>
				<td><%=rsProductInfo("qty")%></td>
				<td><%=rsProductInfo("gauge")%></td>
				<td><%=rsProductInfo("length")%></td>
				<td><%=rsProductInfo("DateLastPurchased")%></td>
				<td><%=rsProductInfo("productdetail1")%></td>
			</tr>
			<%rsProductInfo.MoveNext%>	
		<%Wend%>		
	</table>
<%End If%>	
</div>	
<% 
Set rsSales = Nothing
Set rsRestock = Nothing
Set rsProductInfo = Nothing
DataConn.Close() 
%>
<%
Function GetMonth(month)
	Select Case month
		Case "01"
			GetMonth = "Jan"
		Case "02"
			GetMonth = "Feb"
		Case "03"
			GetMonth = "Mar"
		Case "04"
			GetMonth = "Apr"
		Case "05"
			GetMonth = "May"
		Case "06"
			GetMonth = "Jun"
		Case "07"
			GetMonth = "Jul"
		Case "08"
			GetMonth = "Aug"
		Case "09"
			GetMonth = "Sep"
		Case "10"
			GetMonth = "Oct"			
		Case "11"
			GetMonth = "Nov"
		Case "12"
			GetMonth = "Dec"		
	End Select
End Function
%>
<script>
var ctxL = document.getElementById("lineChart_<%=detailID%>").getContext('2d');
var myLineChart = new Chart(ctxL, {
	type: 'line',
	data: {
		labels: ["<%=month(1)%>", "<%=month(2)%>", "<%=month(3)%>", "<%=month(4)%>", "<%=month(5)%>", "<%=month(6)%>", "<%=month(7)%>", "<%=month(8)%>", "<%=month(9)%>", "<%=month(10)%>", "<%=month(11)%>", "<%=month(12)%>"],
		datasets: [{
			label: "Total Sales: <%=total_sales%>",
			data: [<%=sales(1)%>, <%=sales(2)%>, <%=sales(3)%>, <%=sales(4)%>, <%=sales(5)%>, <%=sales(6)%>, <%=sales(7)%>, <%=sales(8)%>, <%=sales(9)%>, <%=sales(10)%>, <%=sales(11)%>, <%=sales(12)%>],
			backgroundColor: [
			'rgba(105, 0, 132, .2)',
			],
			borderColor: [
			'rgba(200, 99, 132, .7)',
			],
			pointBackgroundColor: 'rgba(200, 99, 132, .7)',
			borderWidth: 2
		},
{
			label: "Total Restock: <%=total_restock%>",
			data: [<%=restock(1)%>, <%=restock(2)%>, <%=restock(3)%>, <%=restock(4)%>, <%=restock(5)%>, <%=restock(6)%>, <%=restock(7)%>, <%=restock(8)%>, <%=restock(9)%>, <%=restock(10)%>, <%=restock(11)%>, <%=restock(12)%>],
			backgroundColor: [
			'rgba(255, 193, 7, .2)',
			],
			borderColor: [
			'rgba(255, 183, 7, .7)',
			],
			pointBackgroundColor: 'rgba(255, 183, 7, .7)',
			borderWidth: 2
		}		
		]
	},
	options: {  
		responsive: true,
		scales: {
			yAxes: [{
				ticks: {
				  beginAtZero: true,
				  callback: function(value) {if (value % 1 === 0) {return value;}}
				}
			}]
		},
		tooltips: {
            callbacks: {
                label: function (tooltipItem, data) {
					let label = data.labels[tooltipItem.index];
                    let value = data.datasets[tooltipItem.datasetIndex].data[tooltipItem.index];
                    return ' ' + label + '.: ' + value;
                }
            }
        }		
	}

});
</script>