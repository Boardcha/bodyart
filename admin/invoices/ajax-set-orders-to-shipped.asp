<%@LANGUAGE="VBSCRIPT"%>
<% Server.ScriptTimeout=300
 %>
<!--#include virtual="/Connections/bodyartforms_sql_ADMIN.asp" -->
<!--#include virtual="/emails/function-send-email.asp"-->
<!--#include virtual="/functions/asp-json.asp"-->
<!--#include virtual="/functions/date-to-iso.asp"-->
<!--#include virtual="/Connections/afterpay-credentials.asp"-->
<%
mailer_type = "order-shipment-notification"
var_tracking = ""

set rsGetInvoice = Server.CreateObject("ADODB.Recordset")
rsGetInvoice.ActiveConnection = DataConn
rsGetInvoice.Source = "SELECT ship_code, shipped, estimated_delivery_date, ID, customer_first, customer_last, email, date_sent, UPS_tracking, USPS_tracking, shipping_type, pay_method, transactionID FROM sent_items WHERE ship_code = 'paid' AND (Review_OrderError <> 1 OR  Review_OrderError IS NULL) AND (shipped = 'Pending shipment') ORDER BY ID DESC"
rsGetInvoice.CursorLocation = 3 'adUseClient
rsGetInvoice.LockType = 1 'Read-only records
rsGetInvoice.Open()

WHILE NOT rsGetInvoice.EOF 

var_email = rsGetInvoice.Fields.Item("email").Value
var_first = rsGetInvoice.Fields.Item("customer_first").Value
var_invoiceid = rsGetInvoice.Fields.Item("ID").Value
var_shipping_type = rsGetInvoice.Fields.Item("shipping_type").Value

	if instr(rsGetInvoice.Fields.Item("shipping_type").Value, "DHL") > 0 then
		var_tracking = "Your tracking # is <strong>" & rsGetInvoice.Fields.Item("USPS_tracking").Value & "</strong>. If you have an account on our website, you can track your package by going to your order history and pressing the Track Order button. Or, you can track your package by going directly to <a href=""https://bodyartforms.com/dhl-tracker.asp?tracking=" & rsGetInvoice.Fields.Item("USPS_tracking").Value & """>this link</a>."
	else
		var_tracking = "Your tracking # is <strong>" & rsGetInvoice.Fields.Item("USPS_tracking").Value & "</strong>. If you have an account on our website, you can track your package by going to your order history and pressing the Track Order button. Or, you can track your package by going directly to <a href=""https://www.usps.com/manage/welcome.htm"">USPS.com</a>"
	end if
	if instr(rsGetInvoice.Fields.Item("shipping_type").Value, "UPS") then
		var_tracking = "Your tracking # is <strong>" & rsGetInvoice.Fields.Item("UPS_tracking").Value & "</strong>. If you have an account on our website, you can track your package by going to your order history and pressing the Track Order button. Or, you can track your package by going directly to <a href=""https://www.ups.com/tracking.html"">UPS.com</a>"
	end if
	if Not IsNull(rsGetInvoice("estimated_delivery_date")) AND rsGetInvoice("estimated_delivery_date") <> "" Then 
		var_tracking = var_tracking & "<br>Estimated delivery date: " & rsGetInvoice("estimated_delivery_date")
	end if
	
	set commUpdate = Server.CreateObject("ADODB.Command")
	commUpdate.ActiveConnection = MM_bodyartforms_sql_STRING
	commUpdate.CommandText = "UPDATE sent_items SET shipped = 'Shipped', date_sent = '"& date() &"'  WHERE ID = " & rsGetInvoice.Fields.Item("ID").Value & " AND ship_code = 'paid'" 
	commUpdate.Execute()

	'========== SENDS TRACKING INFORMATION TO AFTERPAY ==================
	if rsGetInvoice.Fields.Item("pay_method").Value = "Afterpay" then
		Set objAfterPayUpdateTracking = Server.CreateObject("Msxml2.ServerXMLHTTP.6.0")
		objAfterPayUpdateTracking.open "PUT", afterpay_url & "/payments/" & rsGetInvoice.Fields.Item("transactionID").Value & "/courier", false
		objAfterPayUpdateTracking.SetRequestHeader "Authorization", "Basic " & afterpay_api_credential & ""
		objAfterPayUpdateTracking.setRequestHeader "Accept", "application/json"
		objAfterPayUpdateTracking.setRequestHeader "Content-Type", "application/json"
		objAfterPayUpdateTracking.Send("{" & _
				"""shippedAt"": """ & iso8601Date(now()) & """," & _
				"""name"": """ & rsGetInvoice.Fields.Item("shipping_type").Value & """," & _
				"""tracking"": """ & rsGetInvoice.Fields.Item("USPS_tracking").Value & """" & _
			"}")
		
		jsonCapturestring  = objAfterPayUpdateTracking.responseText
		Set oJSON = New aspJSON
		oJSON.loadJSON(jsonCapturestring)
		
		'response.write jsonCapturestring
	end if
	'========== ENDS SENDING TRACKING INFORMATION TO AFTERPAY ==================
%>
<!--#include virtual="/emails/email_variables.asp"-->
<%
rsGetInvoice.movenext()
WEND

rsGetInvoice.Close()
%>
