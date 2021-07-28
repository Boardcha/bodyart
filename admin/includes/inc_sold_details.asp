<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" %>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
if request.querystring("detailid") <> "" then ' check to see if a detailid is provided and if not, just update the main products table	

	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT TBL_OrderSummary.qty, TBL_OrderSummary.DetailID, sent_items.ship_code, sent_items.shipped, sent_items.date_order_placed FROM sent_items INNER JOIN                       TBL_OrderSummary ON sent_items.ID = TBL_OrderSummary.InvoiceID WHERE (sent_items.shipped = N'Waiting for PayPal eCheck to clear' OR sent_items.shipped = N'Review' OR sent_items.shipped = N'PREORDER-REVIEW' OR sent_items.shipped = N'PREORDER-APPROVED' OR sent_items.shipped = N'ON ORDER' OR                      sent_items.shipped = N'ON HOLD' OR sent_items.shipped = N'Pending...') AND (TBL_OrderSummary.DetailID = ? AND sent_items.ship_code = N'paid')"
	'  AND (sent_items.date_order_placed > { fn NOW() } - 1)
	objCmd.Parameters.Append(objCmd.CreateParameter("detailid",3,1,10,request.querystring("detailid")))
	Set rsResearch = objCmd.Execute()

end if

	total_qty = 0
While NOT rsResearch.EOF
	total_qty = total_qty  + rsResearch.Fields.Item("qty").Value
%>
<%
  rsResearch.MoveNext()
Wend 
%>
<p>
<strong>Pending shipment:</strong> <%= total_qty %><br/>
</p>
<%
DataConn.Close()
Set rsSoldInfo = Nothing
%>