<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" %>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<!--#include virtual="/Connections/authnet.asp" -->
<!--#include virtual="/Connections/taxjar.asp"-->
<!--#include virtual="/taxjar/taxjar-nexus-values.asp"-->
<%
' ====== SET VARIABLES =========
invoiceid = request.form("invoiceid")

Set objCmd = Server.CreateObject ("ADODB.Command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT * FROM sent_items WHERE ID = ?" 
objCmd.Parameters.Append(objCmd.CreateParameter("invoiceid",3,1,15, invoiceid))
Set rsGetOrder = objCmd.Execute()

if rsGetOrder.Fields.Item("total_preferred_discount").Value > 0 OR rsGetOrder.Fields.Item("total_coupon_discount").Value > 0 then
    coupon_code = rsGetOrder.Fields.Item("coupon_code").Value
else
    coupon_code = "No Code"
end if

Set objCmd = Server.CreateObject ("ADODB.Command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT DiscountPercent FROM TBLDiscounts WHERE DiscountCode = ?" 
objCmd.Parameters.Append(objCmd.CreateParameter("coupon_code",200,1,50, coupon_code))
Set rsGetCouponDiscount = objCmd.Execute()

Set objCmd = Server.CreateObject("ADODB.Command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT * FROM TBL_OrderSummary INNER JOIN  jewelry ON TBL_OrderSummary.ProductID = jewelry.ProductID WHERE InvoiceID = ?"
objCmd.Parameters.Append(objCmd.CreateParameter("invoiceid",3,1,15, invoiceid))
Set rsGetOrderItems = objCmd.Execute()

if rsGetOrder.Fields.Item("total_gift_cert").Value > 0 then
    Set objCmd = Server.CreateObject ("ADODB.Command")
    objCmd.ActiveConnection = DataConn
    objCmd.CommandText = "SELECT TBLcredits.invoice, TBLcredits.code FROM TBL_Credits_UsedOn INNER JOIN TBLcredits ON TBL_Credits_UsedOn.OriginalCreditID = TBLcredits.ID WHERE TBL_Credits_UsedOn.InvoiceUsedOn = ?" 
    objCmd.Parameters.Append(objCmd.CreateParameter("invoiceid",3,1,15,invoiceid))  
    Set rsGetGiftCertInfo = objCmd.Execute()

    if not rsGetGiftCertInfo.eof then
        gift_cert_code = rsGetGiftCertInfo.Fields.Item("code").Value
        gift_cert_invoice = rsGetGiftCertInfo.Fields.Item("invoice").Value
    end if
end if

' ====== SET VARIABLES =========
db_store_credit = rsGetOrder.Fields.Item("total_store_credit").Value
db_gift_cert = rsGetOrder.Fields.Item("total_gift_cert").Value
db_transactionid = rsGetOrder.Fields.Item("transactionID").Value
db_free_use_now_credits = rsGetOrder.Fields.Item("total_free_credits").Value

LineItem = 0
var_subtotal = 0
preorder_subtotal = 0
coupon_discount = 0
sales_tax = 0
additional_amount = 0
store_credit_refund_due = 0
gift_cert_refund_due = 0
cc_refund_due = 0
authnet_settleAmount = 0
color_addon_fee = 0

' ===== Get order authorized total from authorize.net 
if db_transactionid <> "" then
    strReq = "<?xml version=""1.0"" encoding=""utf-8""?>" _
    & "<getTransactionDetailsRequest xmlns=""AnetApi/xml/v1/schema/AnetApiSchema.xsd"">" _
    & MerchantAuthentication() _
    & "<transId>" & db_transactionid & "</transId>" _
    & "</getTransactionDetailsRequest>"

    Set objGetTransactionDetails = SendApiRequest(strReq)

    ' If succcess retrieve transaction information
    If IsApiResponseSuccess(objGetTransactionDetails) Then
        authnet_settleAmount = objGetTransactionDetails.selectSingleNode("/*/api:transaction/api:settleAmount").Text        
    Else ' if there's an error getting a transaction
    Response.Write "The operation failed with the following errors:<br>" & vbCrLf
    PrintErrors(objGetTransactionDetails)
    End If
end if ' ---- if db_transactionid is found

if request.form("refund-shipping") = 1 then
    shipping_rate = rsGetOrder.Fields.Item("shipping_rate").Value
else
    shipping_rate = 0
end if

if request.form("additional_amount") <> "" then
    additional_amount = request.form("additional_amount")
end if

' ==== Get subtotal ======
    While NOT rsGetOrderItems.EOF 
        For Each item In Request.Form
        If IsNumeric(item) Then
            ' --- if form name is integer then get pricing from table for it    
            If Clng(item) = CLng(rsGetOrderItems.Fields.Item("OrderDetailID").Value) Then
                'Response.Write "MATCHED Form item: " & item & ", Qty: " & Request.Form(item) & vbCrLf 

                LineItem = rsGetOrderItems.Fields.Item("item_price").Value * request.form(item)
                var_subtotal = var_subtotal + LineItem

                color_addon_fee = color_addon_fee + (rsGetOrderItems("qty") * rsGetOrderItems("anodization_fee"))

                if rsGetOrderItems.Fields.Item("customorder").Value = "yes" then
                    preorder_subtotal = (rsGetOrderItems.Fields.Item("item_price").Value * request.form(item)) + preorder_subtotal 
                end if 
            end if
        end if
        Next

    rsGetOrderItems.MoveNext()
    Wend

    ' -------- look up the coupon code and get the coupon discount
    if NOT rsGetCouponDiscount.eof then          
        coupon_discount = (var_subtotal) * (rsGetCouponDiscount.Fields.Item("DiscountPercent").Value/100)
    else
        coupon_discount = 0
    end if


    ' ====== custom order restocking fee, if selected
    if request.form("preorder-restock-fee") = "on" then
        preorder_restock_fee = preorder_subtotal * .15
        preorder_restock_json = formatnumber(preorder_subtotal) * .15
    else
        preorder_restock_fee = 0
        preorder_restock_json = 0
    end if

    'response.write "<br/>restock fee: " & (formatnumber(var_subtotal) * preorder_restock_fee)
    'response.write "<br/>base_subtotal fee: " & formatnumber(var_subtotal)
    'response.write "<br/>preorder_subtotal: " & formatnumber(preorder_subtotal)

    ' ==== Reset subtotal variable
    subtotal_less_discounts = formatnumber(var_subtotal) - formatnumber(preorder_restock_fee) - formatnumber(coupon_discount) - formatnumber(db_free_use_now_credits)

    'subtotal_less_discounts = (formatnumber(var_subtotal) * preorder_restock_fee) - formatnumber(coupon_discount) - formatnumber(db_free_use_now_credits)'

  
    ' UPDATE TAX
    if rsGetOrder.Fields.Item("total_sales_tax").Value > 0 then

    
	if rsGetOrder.Fields.Item("country").Value = "USA" OR rsGetOrder.Fields.Item("country") = "United States" then
        taxjar_to_country = "US"
    end if
    if rsGetOrder.Fields.Item("country") = "Great Britain" OR rsGetOrder.Fields.Item("country") = "Great Britain and Northern Ireland" OR rsGetOrder.Fields.Item("country") = "United Kingdom" then
        taxjar_to_country = "GB"
    end if
    
            Set HttpReq = Server.CreateObject("MSXML2.ServerXMLHTTP")
            HttpReq.open "POST", taxjar_url, false
            HttpReq.setRequestHeader "Content-Type", "application/json"
            HttpReq.SetRequestHeader "Authorization", "Bearer " & taxjar_authorization & ""
            HttpReq.Send("{" & _
                """to_country"":""" & taxjar_to_country & """," & _
                """to_state"":""" & rsGetOrder.Fields.Item("state").Value & """," & _
                """to_zip"":""" & rsGetOrder.Fields.Item("zip").Value & """," & _
                """to_street"": """ & rsGetOrder.Fields.Item("address").Value & """," & _
                """from_country"":""US""," & _
                """from_state"":""TX""," & _
                """from_city"":""Georgetown""," & _
                """from_zip"":""78626""," & _
                """from_street"": ""1966 South Austin Avenue""," & _
                """shipping"":""0""," & _
                """amount"":""" & subtotal_less_discounts & """," & _
                """line_items"": [{" & _
                    """id"":""1""," & _
                    """quantity"": 1," & _
                    """unit_price"": " & subtotal_less_discounts & "," & _
                    """discount"": 0" & _
                "}]," & _
                """nexus_addresses"": [" & _
				    taxjar_nexus_values & _
			    "]" & _
                "}")
    
            'response.write HttpReq.responseText
    
            response_cleaned = HttpReq.responseText
            Dim regEx
                Set regEx = New RegExp
                regEx.Global = true
                regEx.IgnoreCase = True
                regEx.Pattern = "[^A-Za-z0-9,_:.]"
                response_cleaned = regEx.Replace(response_cleaned, "")
    
                response_cleaned = replace(response_cleaned,"tax:", "")
                response_cleaned = replace(response_cleaned,"breakdown:", "")
                response_cleaned = replace(response_cleaned,"jurisdictions:", "")
    
            tax_array = Split(response_cleaned, ",")
            for each x in tax_array
    
                    if instr(x,"amount_to_collect") > 0 then
                        sales_tax = Split(x, ":")(1)
                    end if
            next
            set HttpReq = Nothing

    end if ' if order had tax
    

    ' ====== Find out new total
    subtotal_plus_shipping_and_salestax = subtotal_less_discounts + sales_tax + shipping_rate + additional_amount + color_addon_fee
    ' ========= CALCULATE whether refund goes all to auth.net or some goes to store credit / gift cert
    if CCur(subtotal_plus_shipping_and_salestax) <= CCur(authnet_settleAmount) then
        cc_refund_due = subtotal_plus_shipping_and_salestax

    else
        cc_refund_due = authnet_settleAmount
        ' ==== If a store credit was used ====
        remaining_due = formatnumber(subtotal_plus_shipping_and_salestax) - formatnumber(authnet_settleAmount)
        if CCur(remaining_due) <= CCur(db_store_credit) then
            store_credit_refund_due = remaining_due
        else
            store_credit_refund_due = db_store_credit

            ' ==== If a gift certificate was used ====
            remaining_due = formatnumber(subtotal_plus_shipping_and_salestax) - formatnumber(authnet_settleAmount) - formatnumber(store_credit_refund_due)
            if CCur(remaining_due) <= CCur(db_gift_cert) then
                gift_cert_refund_due = remaining_due
            else
                gift_cert_refund_due = db_gift_cert
            end if ' ===== end gift cert

        end if ' ==== end store credit
    end if ' ====== end calc refund

    ' === sub out credit card refund for store credit if requested
    if request.form("store-credit-only") = "on" then
        store_credit_refund_due = store_credit_refund_due + cc_refund_due 
        cc_refund_due = 0
    end if
%>
{
    "base_subtotal": <%= formatnumber(var_subtotal) %>,
    "subtotal": <%= replace(formatnumber(subtotal_less_discounts), ",", "") %>,
    "preorder_restock_fee": <%= formatnumber(preorder_restock_json) %>,
    "coupon_discount": <%= formatnumber(coupon_discount) %>,
    "sales_tax": <%= formatnumber(sales_tax) %>,
    "shipping_rate": <%= formatnumber(shipping_rate) %>,
    "subtotal_plus_shipping_and_salestax": <%= replace(formatnumber(subtotal_plus_shipping_and_salestax), ",", "") %>,
    "additional_amount": <%= formatnumber(additional_amount) %>,
    "color_addon_fee": <%= formatnumber(color_addon_fee) %>,
    "db_store_credit": <%= formatnumber(db_store_credit) %>,
    "store_credit_refund_due": <%= formatnumber(store_credit_refund_due) %>,
    "gift_cert_refund_due": <%= formatnumber(gift_cert_refund_due) %>,
    "db_gift_cert": <%= formatnumber(db_gift_cert) %>,
    "db_free_use_now_credits": <%= formatnumber(db_free_use_now_credits) %>,
    "gift_cert_code": "<%= gift_cert_code %>",
    "gift_cert_invoice": "<%= gift_cert_invoice %>", 
    "cc_refund_due": <%= replace(formatnumber(cc_refund_due), ",", "") %>,
    "authnet_settleAmount": <%= formatnumber(authnet_settleAmount) %>
}
<%
DataConn.Close()
%>