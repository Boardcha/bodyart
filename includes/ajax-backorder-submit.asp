<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include virtual="/Connections/bodyartforms_sql_ADMIN.asp" -->
<!--#include virtual="/emails/function-send-email.asp"-->

<%
'====== SINCE THIS FILE IS IN ROOT DIRECTORY, MAKE SURE THAT USER IS LOGGED IN VIA ADMIN IN ORDER TO ACCESS CODE ON THIS page

if request.cookies("adminuser") = "yes" AND  request.form("orderdetailid") <> "" then

orderdetailid = request.form("orderdetailid")
Set objCmd = Server.CreateObject ("ADODB.Command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT InvoiceID, ProductID, Customer_ID, DetailID, title, ProductDetail1, Gauge, Length, stock_qty, OrderDetailID, email, customer_first, title, qty, stock_qty, ProductDetail1, ProductDetailID, item_price, PreOrder_Desc, picture, free, type FROM dbo.QRY_OrderDetails WHERE OrderDetailID = ?" 
objCmd.Parameters.Append(objCmd.CreateParameter("orderdetailid",3,1,20, orderdetailid))
Set rsGetInfo = objCmd.Execute()

'================================================================================================
' START store details into a dynamic multidimensional array
reDim array_details_2(12,0)

    array_gauge = ""
    if rsGetInfo("Gauge") <> "" then
        array_gauge = Server.HTMLEncode(rsGetInfo("Gauge"))
    end if
    
    array_length = ""
    if rsGetInfo("Length") <> "" then
        array_length = Server.HTMLEncode(rsGetInfo("Length"))
    end if
    
    array_detail = ""
    if rsGetInfo("ProductDetail1") <> "" then
        array_detail = Server.HTMLEncode(rsGetInfo("ProductDetail1"))
    end if
    
    array_add_new = uBound(array_details_2,2) 
    REDIM PRESERVE array_details_2(12,array_add_new+1) 

    array_details_2(0,array_add_new) = rsGetInfo("ProductDetailID")
    array_details_2(1,array_add_new) = rsGetInfo("qty")
    array_details_2(2,array_add_new) = rsGetInfo("title") 
    array_details_2(3,array_add_new) = array_gauge
    array_details_2(4,array_add_new) = FormatNumber(rsGetInfo("item_price"), -1, -2, -2, -2)
    
    var_preorder_text = ""
    if rsGetInfo("PreOrder_Desc") <> "" then
        var_preorder_text = replace(rsGetInfo("PreOrder_Desc"),"{}", "   ")
    end if
    
    array_details_2(5,array_add_new) = var_preorder_text
    array_details_2(6,array_add_new) = rsGetInfo("ProductID")
    array_details_2(7,array_add_new) = "" '=== item notes
    array_details_2(8,array_add_new) = "" '=== anodization fee
    array_details_2(9,array_add_new)= rsGetInfo("picture")
    array_details_2(10,array_add_new) = array_length
    array_details_2(11,array_add_new) = array_detail
    array_details_2(12,array_add_new) = rsGetInfo("free") 
    
'================================================================================================
' END store details into a dynamic multidimensional array

productdetailid = rsGetInfo.Fields.Item("DetailID").Value
var_customer_name = rsGetInfo.Fields.Item("customer_first").Value
var_customer_email = rsGetInfo.Fields.Item("email").Value
var_invoice_number = rsGetInfo.Fields.Item("InvoiceID").Value
var_customer_number = rsGetInfo.Fields.Item("Customer_ID").Value
var_jewelry_status = rsGetInfo("type")
var_bo_reason = Request.Form("bo_reason")
If var_bo_reason <> "" Then param_bo_reason = ", reason_for_backorder = '" + var_bo_reason + "'"

' Set item to backorder status (and not on review)
set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "UPDATE TBL_OrderSummary SET backorder = 1, backorder_tracking = 1, BackorderReview = 'N'" & param_bo_reason & ", archive_bo_checked_by_who = ? WHERE OrderDetailID = ?"
objCmd.Parameters.Append(objCmd.CreateParameter("archive_bo_checked_by_who",200,1,50, user_name ))
objCmd.Parameters.Append(objCmd.CreateParameter("orderdetailid",3,1,20, orderdetailid))
objCmd.Execute()

' Update quantities on item according to selected drop-down
set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "UPDATE ProductDetails SET qty = ? WHERE ProductDetailID = ?" 
objCmd.Parameters.Append(objCmd.CreateParameter("qty",3,1,20, request.form("bo_qty")))
objCmd.Parameters.Append(objCmd.CreateParameter("productdetailid",3,1,20,productdetailid))
objCmd.Execute()

'Write info to edits log	
set objCmd = Server.CreateObject("ADODB.Command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "INSERT INTO tbl_edits_log (user_id, detail_id, description, edit_date) VALUES (" & user_id & ", ? , ? ,'" & now() & "')"
objCmd.Parameters.Append(objCmd.CreateParameter("qty",3,1,20, rsGetInfo("DetailID") ))
objCmd.Parameters.Append(objCmd.CreateParameter("description",200,1,250, "Automated - Updated qty from " & rsGetInfo("stock_qty") & " to " & request.form("bo_qty") & " - backorder submit page" ))
objCmd.Execute()

' CALCULATE CORRECT PRICE FOR BACKORDERED ITEMS AFTER SALE TO REFUND FOR
set objCmd = Server.CreateObject("ADODB.Command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT TOP (100) PERCENT sent_items.ID, sent_items.coupon_code, sent_items.combined_tax_rate, TBL_OrderSummary.ErrorReportDate, TBL_OrderSummary.ErrorDescription,  sent_items.ship_code, TBL_OrderSummary.qty, ProductDetails.qty AS 'qty_instock', TBL_OrderSummary.item_price, ProductDetails.ProductDetail1, ProductDetails.location, ProductDetails.Gauge, ProductDetails.Length, jewelry.title, ProductDetails.ProductDetailID, ProductDetails.BinNumber_Detail, TBL_OrderSummary.OrderDetailID, TBL_OrderSummary.ProductID, TBL_OrderSummary.item_problem, TBL_OrderSummary.ErrorQtyMissing,  (jewelry.title + ' ' + ISNULL(ProductDetails.Gauge, '') + ' ' + ISNULL(ProductDetails.Length, '') + ' ' + ISNULL(ProductDetails.ProductDetail1, '')) as description FROM sent_items INNER JOIN TBL_OrderSummary ON sent_items.ID = TBL_OrderSummary.InvoiceID INNER JOIN ProductDetails ON TBL_OrderSummary.DetailID = ProductDetails.ProductDetailID INNER JOIN jewelry ON TBL_OrderSummary.ProductID = jewelry.ProductID WHERE TBL_OrderSummary.backorder = 1 AND ID = ? ORDER BY sent_items.ID"
objCmd.Parameters.Append(objCmd.CreateParameter("invoiceid",3,1,12, var_invoice_number))
set rsGetItems = Server.CreateObject("ADODB.Recordset")
rsGetItems.CursorLocation = 3 'adUseClient
rsGetItems.Open objCmd

If NOT rsGetItems.EOF Then
	'==============  GET COUPON DISCOUNT / IF ANY ============================================
	set objCmd = Server.CreateObject("ADODB.Command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT DiscountPercent FROM TBLDiscounts WHERE DiscountCode = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("coupon_code",200,1,50,rsGetItems.Fields.Item("coupon_code").Value))
	Set rsGetCouponDiscount = objCmd.Execute()
End If

While NOT rsGetItems.EOF 
	
	If NOT rsGetCouponDiscount.eof then
		var_item_price = FormatNumber((rsGetItems.Fields.Item("item_price").Value - ((rsGetCouponDiscount.Fields.Item("DiscountPercent").Value / 100) * rsGetItems.Fields.Item("item_price").Value)) * rsGetItems.Fields.Item("ErrorQtyMissing").Value, -1, -2, -0, -2)                        
	Else
		var_item_price = FormatNumber(rsGetItems.Fields.Item("item_price").Value * rsGetItems.Fields.Item("qty").Value, -1, -2, -0, -2)
	End if

	' Add on tax to refund 
	If rsGetItems.Fields.Item("combined_tax_rate").Value > 0 then
		var_item_price = var_item_price + (var_item_price * rsGetItems.Fields.Item("combined_tax_rate").Value)
	End if
	var_refund_total = FormatNumber(Ccur(var_refund_total) + ccur(var_item_price), -1, -2, -0, -2)
	rsGetItems.MoveNext
Wend

If var_refund_total > 0 then

	Set objCrypt = Server.CreateObject("Bodyartforms.BAFCrypt")
	password = "3uBRUbrat77V"
	data = var_invoice_number & "|" & var_refund_total & "|" & var_customer_number
	encrypted_code = objCrypt.Encrypt(password, data)

	
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "DELETE FROM TBL_Refunds_backordered_items WHERE invoice_id = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("invoice_id",3,1,15, var_invoice_number))
	objCmd.Execute()

	
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "INSERT INTO TBL_Refunds_backordered_items (invoice_id, refund_total, encrypted_code) VALUES (?,?,?)"
	objCmd.Parameters.Append(objCmd.CreateParameter("invoice_id",3,1,15, var_invoice_number))
	objCmd.Parameters.Append(objCmd.CreateParameter("refund_total",6,1,20, var_refund_total))
	objCmd.Parameters.Append(objCmd.CreateParameter("encrypted_code",200,1,250, encrypted_code))
	objCmd.Execute()

	Set objCrypt = Nothing
	Set objCrypt = Server.CreateObject("Bodyartforms.BAFCrypt")
	password = "3uBRUbrat77V"
	data = encrypted_code
	decrypted = objCrypt.Decrypt(password, data)
	response.write "decrypted: " & decrypted
	Set objCrypt = Nothing

End if
		
Set objCmd = Nothing
mailer_type = "backorder"
%>
<!--#include virtual="/checkout/inc_random_code_generator.asp"-->
<!--#include virtual="/includes/inc-dupe-onetime-codes.asp"--> 
<%
'================ Prepare a one time use coupon for the backorder hassle
var_cert_code = getPassword(15, extraChars, firstNumber, firstLower, firstUpper, firstOther, latterNumber, latterLower, latterUpper, latterOther)

' Call function
var_cert_code = CheckDupe(var_cert_code)

'======= Store one time coupon code
set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "INSERT INTO TBLDiscounts (DiscountCode, DateExpired, coupon_single_email, DiscountPercent, coupon_single_use, DateAdded, DiscountType, active, dateactive, coupon_assigned, DiscountDescription) VALUES (?, GETDATE()+730, ?, 15, 1, GETDATE(), 'Percentage', 'A', GETDATE()-1, 1, 'Backordered item discount')"
objCmd.Parameters.Append(objCmd.CreateParameter("Code",200,1,30,var_cert_code ))
objCmd.Parameters.Append(objCmd.CreateParameter("Email",200,1,30, var_customer_email ))
objCmd.Execute()
%>
<!--#include virtual="/emails/email_variables.asp"-->
<%

response.write "LOGGED IN"
else
response.write "NOT LOGGED IN"
end if '===== user_name <> ""

DataConn.Close()
Set rsGetInfo = Nothing
%>