<%@LANGUAGE="VBSCRIPT"%>
<% 
Server.ScriptTimeout=5000
 %>
<!--#include virtual="/Connections/bodyartforms_sql_ADMIN.asp" -->
<!--#include virtual="/emails/function-send-email.asp"-->
<!--#include virtual="/functions/asp-json.asp"-->
<!--#include virtual="/functions/date-to-iso.asp"-->
<!--#include virtual="/Connections/afterpay-credentials.asp"-->
<%
mailer_type = "order-shipment-notification"
done_mailing_certs = "yes" ' no purpose served on this page other than making it working with file /emails/email_variables.asp

set objCmd = Server.CreateObject("ADODB.Command")
objCmd.ActiveConnection = DataConn

if request.form("resend_email") = "" then '======= from /admin/invoice.asp page
	objCmd.CommandText = "SELECT * FROM sent_items WHERE ship_code = 'paid' AND (Review_OrderError <> 1 OR  Review_OrderError IS NULL) AND (shipped = 'Pending shipment') ORDER BY ID DESC"
else
	objCmd.CommandText = "SELECT * FROM sent_items WHERE ID = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("invoiceid",3,1,15, request.form("resend_email")  ))
end if
set rsGetInvoice = Server.CreateObject("ADODB.Recordset")
rsGetInvoice.CursorLocation = 3 'adUseClient
rsGetInvoice.LockType = 1 'Read-only records
rsGetInvoice.Open objCmd

WHILE NOT rsGetInvoice.EOF 

	'======= UPDATE ORDER STATUS ==========================
	set objCmd = Server.CreateObject("ADODB.Command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "UPDATE sent_items SET shipped = 'Shipped', date_sent = '"& date() &"'  WHERE ID = " & rsGetInvoice("ID") & " AND ship_code = 'paid'" 
	objCmd.Execute()

	var_email = rsGetInvoice.Fields.Item("email").Value
	var_first = rsGetInvoice.Fields.Item("customer_first").Value
	var_invoiceid = rsGetInvoice.Fields.Item("ID").Value
	var_shipping_type = rsGetInvoice.Fields.Item("shipping_type").Value
	var_tracking = ""
	estimated_delivery = ""

	if Not IsNull(rsGetInvoice("estimated_delivery_date")) AND rsGetInvoice("estimated_delivery_date") <> "" AND (rsGetInvoice("country") = "USA" OR rsGetInvoice("country") = "US") Then 
		estimated_delivery = "Estimated delivery date: " & rsGetInvoice("estimated_delivery_date")
	end if

		var_tracking = "<div style='font-family:Arial;color: #ffffff;;background-color:#696986;padding:20px;border-radius:10px'>Your tracking # is " & rsGetInvoice.Fields.Item("USPS_tracking").Value & "<br>Shipped via " & var_shipping_type & "<br>" & estimated_delivery & "<br><br><a style='font-family:Arial;font-size:16px;color: #ffffff;;background-color:#41415a;padding:10px;font-weight:bold;text-decoration:none' href='"
%>
<!--#include virtual="/admin/packing/tracker-builder.asp"-->
<%
		var_tracking = var_tracking & "'>TRACK YOUR PACKAGE</a><br><br><i>Please note that it can take 1-2 days for tracking information to become available.</i></div>"
	

	'================================================================================================
	' START store details into a dynamic multidimensional array
	reDim array_details_2(12,0)

	set objCmd = Server.CreateObject("ADODB.Command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT TBL_OrderSummary.qty, gauge, length, ProductDetail1, title, ProductDetailID, TBL_OrderSummary.ProductID, picture, PreOrder_Desc, item_price, free  FROM TBL_OrderSummary INNER JOIN jewelry ON TBL_OrderSummary.ProductID = jewelry.ProductID INNER JOIN ProductDetails ON TBL_OrderSummary.DetailID = ProductDetails.ProductDetailID WHERE TBL_OrderSummary.InvoiceID = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("invoiceid",3,1,15, rsGetInvoice("ID")  ))
	set rsGetItems = Server.CreateObject("ADODB.Recordset")
	rsGetItems.CursorLocation = 3 'adUseClient
	rsGetItems.LockType = 1 'Read-only records
	rsGetItems.Open objCmd

	while NOT rsGetItems.EOF

		array_gauge = ""
		if rsGetItems("Gauge") <> "" then
			array_gauge = Server.HTMLEncode(rsGetItems("Gauge"))
		end if
		
		array_length = ""
		if rsGetItems("Length") <> "" then
			array_length = Server.HTMLEncode(rsGetItems("Length"))
		end if
		
		array_detail = ""
		if rsGetItems("ProductDetail1") <> "" then
			array_detail = Server.HTMLEncode(rsGetItems("ProductDetail1"))
		end if
		
		array_add_new = uBound(array_details_2,2) 
		REDIM PRESERVE array_details_2(12,array_add_new+1) 

		array_details_2(0,array_add_new) = rsGetItems("ProductDetailID")
		array_details_2(1,array_add_new) = rsGetItems("qty")
		array_details_2(2,array_add_new) = rsGetItems("title") 
		array_details_2(3,array_add_new) = array_gauge
		array_details_2(4,array_add_new) = FormatNumber(rsGetItems("item_price"), -1, -2, -2, -2)
		
		var_preorder_text = ""
		if rsGetItems("PreOrder_Desc") <> "" then
			var_preorder_text = replace(rsGetItems("PreOrder_Desc"),"{}", "   ")
		end if
		
		array_details_2(5,array_add_new) = var_preorder_text
		array_details_2(6,array_add_new) = rsGetItems("ProductID")
		array_details_2(7,array_add_new) = "" ' item notes
		array_details_2(8,array_add_new) = "" '=== anodization fee
		array_details_2(9,array_add_new)= rsGetItems("picture")
		array_details_2(10,array_add_new) = array_length
		array_details_2(11,array_add_new) = array_detail
		array_details_2(12,array_add_new) = rsGetItems("free") 
		
		rsGetItems.MoveNext()
	Wend


	'================================================================================================
	' END store details into a dynamic multidimensional array
	
%>
<!--#include virtual="/emails/email_variables.asp"-->
<%
response.write "invoice " & rsGetInvoice("ID") & " processed<br>"
rsGetInvoice.movenext()
WEND

rsGetItems.close()
rsGetInvoice.Close()
set rsGetItems = nothing
set rsGetInvoice = nothing
%>
