<%
' =================================================================================
' Update main order totals & other information
' This code is only run if payment is actually completed
' =================================================================================

if preorder_shipping_notice = "yes" then
    var_toggle_preorder = ", preorder = 1, shipped = 'CUSTOM ORDER IN REVIEW'"
end if

set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "UPDATE sent_items SET total_sales_tax = total_sales_tax + ?, taxes_state_only = taxes_state_only + ?, taxes_county_only = taxes_county_only + ?, taxes_city_only = taxes_city_only + ?, taxes_special_only = taxes_special_only + ?, total_gift_cert = total_gift_cert + ?, total_coupon_discount = total_coupon_discount + ?, total_preferred_discount = total_preferred_discount + ?, total_store_credit = total_store_credit + ?, total_free_credits = total_free_credits + ? " & var_toggle_preorder & " WHERE ID = ?"		
    objCmd.Parameters.Append(objCmd.CreateParameter("@total_sales_tax",6,1,10,session("amount_to_collect")))
    objCmd.Parameters.Append(objCmd.CreateParameter("@taxes_state_only",6,1,10,session("state_tax_collectable")))
    objCmd.Parameters.Append(objCmd.CreateParameter("@taxes_county_only",6,1,10,session("county_tax_collectable")))
    objCmd.Parameters.Append(objCmd.CreateParameter("@taxes_city_only",6,1,10,session("city_tax_collectable")))
    objCmd.Parameters.Append(objCmd.CreateParameter("@taxes_special_only",6,1,10,session("special_district_tax_collectable")))
    objCmd.Parameters.Append(objCmd.CreateParameter("@total_gift_cert",6,1,10,FormatNumber(var_total_giftcert_used, -1, -2, -2, -2)))
    objCmd.Parameters.Append(objCmd.CreateParameter("@total_coupon_discount",6,1,10,FormatNumber(var_couponTotal, -1, -2, -2, -2)))
    objCmd.Parameters.Append(objCmd.CreateParameter("@total_preferred_discount",6,1,10,FormatNumber(total_preferred_discount, -1, -2, -2, -2)))
    objCmd.Parameters.Append(objCmd.CreateParameter("@total_store_credit",6,1,10,FormatNumber(session("storeCredit_used"),2)))
    objCmd.Parameters.Append(objCmd.CreateParameter("@total_free_credits",6,1,10,FormatNumber(var_credit_now,2)))

    objCmd.Parameters.Append(objCmd.CreateParameter("invoice_id",3,1,15,session("invoiceid")))
    objCmd.Execute()

    ' ===== if a coupon was used write a note to the order 
    if session("preferred") = "yes" then
        var_addon_coupon_code = "YTG89R57"
    end if
    if Session("CouponCode") <> "" then
        var_addon_coupon_code = Session("CouponCode")
    end if
    if var_addon_coupon_code <> "" then
        set objCmd = Server.CreateObject("ADODB.command")
        objCmd.ActiveConnection = DataConn
        objCmd.CommandText = "INSERT INTO tbl_invoice_notes (user_id, invoice_id, note) VALUES (?,?,?)"
        objCmd.Parameters.Append(objCmd.CreateParameter("user_id",3,1,10,1))
        objCmd.Parameters.Append(objCmd.CreateParameter("invoiceid",3,1,15, session("invoiceid")))
        objCmd.Parameters.Append(objCmd.CreateParameter("note",200,1,250,"Coupon " & var_addon_coupon_code & " used on add-on items"))
        objCmd.Execute()
    end if


' =================================================================================
'Save order details -- array is generated in inc_orderdetails_toarray.asp
' =================================================================================
if request.cookies("OrderAddonsActive") <> "" then
    var_addons_db_flag = 1
else
    var_addons_db_flag = 0
end if


For i = 0 to (ubound(array_details_2, 2) - 1) ' loop through array
    set objCmd = Server.CreateObject("ADODB.command")
    objCmd.ActiveConnection = DataConn
    'objCmd.CommandType = 4
    'objCmd.CommandText = "Proc_Checkout4_InsertOrder"
	
	'If ProductID flagged as "waiting-list", meaning if customer comes from waiting-list email notification, save this info to the "referrer" field.
	If Session(array_details_2(6,i)) = "waiting-list" Then 
		var_referrer = "'waiting-list'" 
	ElseIf Session(array_details_2(13,i)) = 2 Then ' 2 = it is added to cart back from saved items
		var_referrer = "'save-for-later'" 
	Else 
		var_referrer = "NULL"
	End If
    objCmd.CommandText = "INSERT INTO TBL_OrderSummary (InvoiceID, ProductID, DetailID, qty, item_price, notes, PreOrder_Desc, item_wlsl_price, addon_item, referrer) VALUES (?,?,?,?,?,?,?,?, " & var_addons_db_flag & "," & var_referrer & ")"
            objCmd.Parameters.Append(objCmd.CreateParameter("invoiceid",3,1,15,session("invoiceid")))
            objCmd.Parameters.Append(objCmd.CreateParameter("productid",3,1,15,array_details_2(6,i)))
            objCmd.Parameters.Append(objCmd.CreateParameter("detailid",3,1,15,array_details_2(0,i)))
            objCmd.Parameters.Append(objCmd.CreateParameter("qty",3,1,10,array_details_2(1,i)))
            objCmd.Parameters.Append(objCmd.CreateParameter("item_price",6,1,10,array_details_2(4,i)))
            objCmd.Parameters.Append(objCmd.CreateParameter("item_notes",200,1,50,array_details_2(7,i)))
            objCmd.Parameters.Append(objCmd.CreateParameter("preorder_notes",200,1,2000,array_details_2(5,i)))
            objCmd.Parameters.Append(objCmd.CreateParameter("item_wlsl_price",6,1,10,array_details_2(8,i)))
    objCmd.Execute()
next ' loop through array

' =================================================================================

' =================================================================================
' Retrieve the main order information to send email receipt
' =================================================================================
    set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT * FROM sent_items WHERE ID = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("invoice_id",3,1,15,session("invoiceid")))
    set rsGetAddonsMainOrder = objCmd.Execute()
    
    if not rsGetAddonsMainOrder.eof then
        session("email") = rsGetAddonsMainOrder.Fields.Item("email").Value
        session("shipping_first") = rsGetAddonsMainOrder.Fields.Item("customer_first").Value
        session("shipping_last") = rsGetAddonsMainOrder.Fields.Item("customer_last").Value
        session("shipping_company") = rsGetAddonsMainOrder.Fields.Item("company").Value
        session("shipping_address1") = rsGetAddonsMainOrder.Fields.Item("address").Value
        session("shipping_address2") = rsGetAddonsMainOrder.Fields.Item("address2").Value
        session("city") = rsGetAddonsMainOrder.Fields.Item("city").Value
        session("state") = rsGetAddonsMainOrder.Fields.Item("state").Value
        session("shipping_province") = rsGetAddonsMainOrder.Fields.Item("province").Value
        session("shipping_zip") = rsGetAddonsMainOrder.Fields.Item("zip").Value
        session("country") = rsGetAddonsMainOrder.Fields.Item("country").Value
    end if
%>
