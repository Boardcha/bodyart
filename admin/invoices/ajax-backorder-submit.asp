<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<!--#include virtual="emails/function-send-email.asp"-->

<%
orderdetailid = request.form("orderdetailid")
Set objCmd = Server.CreateObject ("ADODB.Command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT InvoiceID, ProductID, DetailID, title, ProductDetail1, Gauge, Length, stock_qty, OrderDetailID, email, customer_first, title, qty, ProductDetail1, ProductDetailID, item_price, PreOrder_Desc, picture, free FROM dbo.QRY_OrderDetails WHERE OrderDetailID = ?" 
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
    array_details_2(7,array_add_new) = "" ' item notes
    array_details_2(8,array_add_new) = 0
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
var_bo_reason = Request.Form("bo_reason")

' Set item to backorder status (and not on review)
set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "UPDATE TBL_OrderSummary SET backorder = 1, backorder_tracking = 1, BackorderReview = 'N' WHERE OrderDetailID = ?"
objCmd.Parameters.Append(objCmd.CreateParameter("orderdetailid",3,1,20, orderdetailid))
objCmd.Execute()

' Update quantities on item according to selected drop-down
set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "UPDATE ProductDetails SET qty = ? WHERE ProductDetailID = ?" 
objCmd.Parameters.Append(objCmd.CreateParameter("qty",3,1,20, request.form("bo_qty")))
objCmd.Parameters.Append(objCmd.CreateParameter("productdetailid",3,1,20,productdetailid))
objCmd.Execute()

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
<!--#include virtual="emails/email_variables.asp"-->
<%

DataConn.Close()
Set rsGetInfo = Nothing
%>