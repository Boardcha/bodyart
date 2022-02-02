<%
Server.ScriptTimeout = 1000
%>
<!--#include virtual="/Connections/sql_connection.asp" -->
<!--#include virtual="/Connections/dhl-auth-v4.asp"-->
<!--#include virtual="/emails/function-send-email.asp"-->
<!--#include virtual="/scheduled-tasks/function-item-builder.asp"-->
<!--#include virtual="/scheduled-tasks/function-get-delivery-status.asp"-->
<%
'=== CHECK ORDERS WILL BE DELIVERED TODAY ===
Set objCmd = Server.CreateObject("ADODB.Command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT ID, customer_first, email, estimated_delivery_date, USPS_tracking, shipping_type, date_order_placed FROM sent_items WHERE estimated_delivery_date = CONVERT(VARCHAR(10), GETDATE(), 23) AND delivering_today_email_sent = 0" 
Set rsGetInvoice = objCmd.Execute()

reDim array_details_2(12,0)

While Not rsGetInvoice.EOF 

	status = getDeliveryStatus(rsGetInvoice("USPS_tracking"))
	var_tracking = ""
	%>
	<!--#include virtual="/admin/packing/tracker-builder.asp"-->
	<%
	If status = "OUT_FOR_DELIVERY" Then 
		GetOrderItems(rsGetInvoice("ID")) 'Calls function that build items array
		mailer_type = "OUT_FOR_DELIVERY"
		var_email = rsGetInvoice("email")
		var_first = rsGetInvoice("customer_first")
		%>
		<!--#include virtual="/emails/email_variables.asp"-->
		<%
		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "UPDATE sent_items SET delivering_today_email_sent = 1 WHERE ID = " & rsGetInvoice("ID")
																										
		objCmd.Execute()	
	End If
	rsGetInvoice.MoveNext
Wend
rsGetInvoice.Close 

Response.Write "Successfuly completed." 
Set rsGetInvoice = Nothing
DataConn.Close()
Set DataConn = Nothing
%>
