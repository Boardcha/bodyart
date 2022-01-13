<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<!--#include virtual="emails/function-send-email.asp"-->
<%
var_agenda = request.form("agenda")
var_invoiceid = request.form("invoiceid")

set objCmd = Server.CreateObject("ADODB.Command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT TOP (100) PERCENT ID, shipped, customer_first, customer_last, email, country, PackagedBy, ship_code, customer_ID, coupon_code, combined_tax_rate FROM sent_items WHERE ID = ?"
objCmd.Parameters.Append(objCmd.CreateParameter("invoiceid",3,1,12, var_invoiceid))
set rsGetInvoice = objCmd.Execute()

If NOT rsGetInvoice.EOF Then

'==============  GET COUPON DISCOUNT / IF ANY ============================================
set objCmd = Server.CreateObject("ADODB.Command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT DiscountPercent FROM TBLDiscounts WHERE DiscountCode = ?"
objCmd.Parameters.Append(objCmd.CreateParameter("coupon_code",200,1,50,rsGetInvoice.Fields.Item("coupon_code").Value))
Set rsGetCouponDiscount = objCmd.Execute()

'================== GET ITEMS FROM ORDER ==============================================
set objCmd = Server.CreateObject("ADODB.Command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT TOP (100) PERCENT sent_items.ID, TBL_OrderSummary.ErrorReportDate, TBL_OrderSummary.ErrorDescription,  sent_items.ship_code, TBL_OrderSummary.qty, ProductDetails.qty AS 'qty_instock', TBL_OrderSummary.item_price, ProductDetails.ProductDetail1, ProductDetails.location, ProductDetails.Gauge, ProductDetails.Length, jewelry.title, ProductDetails.ProductDetailID, ProductDetails.BinNumber_Detail, TBL_OrderSummary.OrderDetailID, TBL_OrderSummary.ProductID, TBL_OrderSummary.item_problem, TBL_OrderSummary.ErrorQtyMissing,  (jewelry.title + ' ' + ISNULL(ProductDetails.Gauge, '') + ' ' + ISNULL(ProductDetails.Length, '') + ' ' + ISNULL(ProductDetails.ProductDetail1, '')) as description FROM sent_items INNER JOIN TBL_OrderSummary ON sent_items.ID = TBL_OrderSummary.InvoiceID INNER JOIN ProductDetails ON TBL_OrderSummary.DetailID = ProductDetails.ProductDetailID INNER JOIN jewelry ON TBL_OrderSummary.ProductID = jewelry.ProductID WHERE TBL_OrderSummary.ErrorOnReview = 1 AND ID = ? ORDER BY sent_items.ID"
objCmd.Parameters.Append(objCmd.CreateParameter("invoiceid",3,1,12, var_invoiceid))

set rsGetItems = Server.CreateObject("ADODB.Recordset")
rsGetItems.CursorLocation = 3 'adUseClient
rsGetItems.Open objCmd

'=== Only want to run this code once on page load, do not place below otherwise it will loop
var_giftcert = "yes"
var_cert_id = 0
var_refund_total = 0
%>
<!--#include virtual="/checkout/inc_random_code_generator.asp"-->
<!--#include virtual="/checkout/inc_giftcert_check_dupes.asp"--> 

<%
var_cert_code = strRandomCode
var_create_neworder = ""

    if var_agenda = "approve" then

        If NOT rsGetItems.EOF Then

            '==== CHECK TO SEE IF A NEW ORDER NEEDS TO BE CREATED (ONLY IF ITEMS WILL BE SHIPPED)======
            While NOT rsGetItems.EOF 
                if (rsGetItems.Fields.Item("qty_instock").Value >= rsGetItems.Fields.Item("ErrorQtyMissing").Value) then
                    var_create_neworder = "yes"
                end if 
            rsGetItems.MoveNext()
            Wend
            rsGetItems.MoveFirst()

            '===== CREATE A NEW EMPTY ORDER ===================================================
            if rsGetInvoice.Fields.Item("country").Value = "USA" then
                var_shipping_type = "DHL Basic mail"
            else
                var_shipping_type = "DHL GlobalMail Packet Priority"
            end if

            if var_create_neworder = "yes" then

            set objCmd = Server.CreateObject("ADODB.Command")
            objCmd.ActiveConnection = DataConn
            objCmd.CommandText = "INSERT INTO sent_items (shipped, customer_ID, customer_first, customer_last, company, address, address2, city, state, province, zip, country, email, date_order_placed, shipping_rate, shipping_type, ship_code, phone, pay_method, UPS_Service, autoclave, reship_processed_by) SELECT 'Pending...', customer_ID, customer_first, customer_last, company, address, address2, city, state, province, zip, country, email, '" & now() & "',0 ,'" & var_shipping_type & "' , 'paid', phone, pay_method, '', autoclave, ? FROM sent_items WHERE ID = ?" 
            objCmd.Parameters.Append(objCmd.CreateParameter("user_name",200,1,50, user_name))
            objCmd.Parameters.Append(objCmd.CreateParameter("invoiceid",3,1,15,var_invoiceid))
            objCmd.Execute() 
            
            Set objCmd = Server.CreateObject ("ADODB.Command")
            objCmd.ActiveConnection = DataConn
            objCmd.CommandText = "SELECT TOP(1) ID FROM sent_items WHERE email = ? ORDER BY ID DESC" 
            objCmd.Parameters.Append(objCmd.CreateParameter("email",200,1,75,rsGetInvoice.Fields.Item("email").Value))
            Set rsGetNewestInvoice = objCmd.Execute()	

            move_to_invoice = rsGetNewestInvoice.Fields.Item("ID").Value

            end if '===== var_create_neworder = "yes"

            While NOT rsGetItems.EOF 
            if rsGetItems.Fields.Item("ErrorQtyMissing").Value > 0 then

                ' ===== IF ITEM IS IN STOCK, THEN SHIP IT OUT AND DEDUCT QUANTITIES ===========
                if (rsGetItems.Fields.Item("qty_instock").Value >= rsGetItems.Fields.Item("ErrorQtyMissing").Value) then

                    ' ===== INSERT REPLACEMENT ITEMS INTO ORDER ===============================
                    set objCmd = Server.CreateObject("ADODB.Command")
                    objCmd.ActiveConnection = DataConn
                    objCmd.CommandText = "INSERT INTO TBL_OrderSummary(InvoiceID, ProductID, DetailID, qty, item_price) SELECT ? , ProductID, DetailID, ErrorQtyMissing, 0 FROM TBL_OrderSummary WHERE OrderDetailID = ?" 
                    objCmd.Parameters.Append(objCmd.CreateParameter("invoiceid",3,1,15,move_to_invoice))
                    objCmd.Parameters.Append(objCmd.CreateParameter("orderdetailid",3,1,15,rsGetItems.Fields.Item("OrderDetailID").Value))
                    objCmd.Execute()

                    ' Deduct inventory
                    set objCmd = Server.CreateObject("ADODB.command")
                    objCmd.ActiveConnection = DataConn
                    objCmd.CommandText = "UPDATE ProductDetails SET qty = qty - " & rsGetItems.Fields.Item("ErrorQtyMissing").Value & ", DateLastPurchased = '"& date() &"' WHERE ProductDetailID = " & rsGetItems.Fields.Item("ProductDetailID").Value
                    objCmd.Execute()

                    '======================= Write info to edits log	
                    set objCmd = Server.CreateObject("ADODB.Command")
                    objCmd.ActiveConnection = DataConn
                    objCmd.CommandText = "INSERT INTO tbl_edits_log (user_id, detail_id, description, edit_date) VALUES (" & user_id & ", " & rsGetItems("ProductDetailID") & ",'Automated - Updated (deducted) qty from " & rsGetItems("qty_instock") & " to " & rsGetItems("qty_instock") - rsGetItems("ErrorQtyMissing") & " - reship items window','" & now() & "')"
                    objCmd.Execute()
                    Set objCmd = Nothing

                    ' ====== Write item notes
                    set objCmd = Server.CreateObject("ADODB.command")
                    objCmd.ActiveConnection = DataConn
                    objCmd.CommandText = "UPDATE TBL_OrderSummary SET notes = '" & rsGetItems.Fields.Item("item_problem").Value & " - See invoice " & move_to_invoice & "' WHERE OrderDetailID = " & rsGetItems.Fields.Item("OrderDetailID").Value
                    objCmd.Execute()

                    email_stocked_items = email_stocked_items & "<li>" & rsGetItems.Fields.Item("description").Value & "</li>"
                    add_reship_notes = "yes"

                else ' ====== ITEM IS NOT IN STOCK ... DO NOT DEDUCT OR WRITE TO ORDER

                    ' ============== CALCULATE CORRECT PRICE AFTER SALE TO REFUND FOR
                    if NOT rsGetCouponDiscount.eof then
                        var_item_price = FormatNumber((rsGetItems.Fields.Item("item_price").Value - ((rsGetCouponDiscount.Fields.Item("DiscountPercent").Value / 100) * rsGetItems.Fields.Item("item_price").Value)) * rsGetItems.Fields.Item("ErrorQtyMissing").Value, -1, -2, -0, -2)                        
                    else
                        var_item_price = FormatNumber(rsGetItems.Fields.Item("item_price").Value * rsGetItems.Fields.Item("ErrorQtyMissing").Value, -1, -2, -0, -2)
                    end if
                    
                    '===== Add on tax to refund ==============
                    if rsGetInvoice.Fields.Item("combined_tax_rate").Value > 0 then
                        var_item_price = var_item_price + (var_item_price * rsGetInvoice.Fields.Item("combined_tax_rate").Value)
                    end if
                    var_refund_total = FormatNumber(Ccur(var_refund_total) + ccur(var_item_price), -1, -2, -0, -2)
                    

                    ' ===== If customer is registered, issue a store credit for out of stock items ===========
                    if rsGetInvoice.Fields.Item("customer_ID").Value > 0 then
                    
                        set objCmd = Server.CreateObject("ADODB.Command")
                        objCmd.ActiveConnection = DataConn
                        objCmd.CommandText = "UPDATE customers SET credits = credits + " & var_refund_total & " WHERE customer_ID = ?"
                        objCmd.Parameters.Append(objCmd.CreateParameter("customerid",3,1,12, rsGetInvoice.Fields.Item("customer_ID").Value))
                        objCmd.Execute()

                        add_storecredit_notes = "yes"

                    ' ===== If customer is NOT registerd, issue a gift certificate for out of stock items ================
                    else '=====================

                        var_cert_code = strRandomCode
        
                        '======= Call function to check for duplicates
                        var_cert_code = CheckDupe(var_cert_code)


                        set objCmd = Server.CreateObject("ADODB.command")
                        objCmd.ActiveConnection = DataConn
                        objCmd.CommandText = "INSERT INTO TBLCredits (invoice, name, rec_name, rec_email, message, code, amount) VALUES (?,?,?,?,?,?,?)"
                        objCmd.Parameters.Append(objCmd.CreateParameter("invoiceid",3,1,10,var_invoiceid))
                        objCmd.Parameters.Append(objCmd.CreateParameter("purchaser_name",200,1,30, rsGetInvoice.Fields.Item("customer_first").Value))
                        objCmd.Parameters.Append(objCmd.CreateParameter("recipient_name",200,1,30,rsGetInvoice.Fields.Item("customer_first").Value))
                        objCmd.Parameters.Append(objCmd.CreateParameter("recipient_email",200,1,50, rsGetInvoice.Fields.Item("email").Value))
                        objCmd.Parameters.Append(objCmd.CreateParameter("message",200,1,250, "Automated setup for reshipping items"))
                        objCmd.Parameters.Append(objCmd.CreateParameter("code",200,1,50,var_cert_code))
                        objCmd.Parameters.Append(objCmd.CreateParameter("amount",6,1,20, var_refund_total))
                        objCmd.Execute()

                        Set objCmd = Server.CreateObject ("ADODB.Command")
                        objCmd.ActiveConnection = DataConn
                        objCmd.CommandText = "SELECT TOP(1) ID FROM TBLCredits WHERE code = ? AND invoice = ? ORDER BY ID DESC" 
                        objCmd.Parameters.Append(objCmd.CreateParameter("code",200,1,75,var_cert_code))
                        objCmd.Parameters.Append(objCmd.CreateParameter("invoiceid",3,1,10,var_invoiceid))
                        Set rsGetGiftCertId = objCmd.Execute()
                        
                        var_cert_id = rsGetGiftCertId.Fields.Item("ID").Value

                        add_giftcert_notes = "yes"

                    end if '==== issue gift cert or store credit

                    email_outofstock_items = email_outofstock_items & "<li>" & rsGetItems.Fields.Item("description").Value & "</li>"
                    
                end if '====== if items are out of stock or not

            end if ' ErrorQtyMissing > 0 
            rsGetItems.MoveNext()
            Wend
        end if '===== NOT rsGetItems.EOF

        ' ====== REMOVE ALL ITEMS FROM REVIEW PAGE ============================================
        set objCmd = Server.CreateObject("ADODB.Command")
        objCmd.ActiveConnection = DataConn
        objCmd.CommandText = "UPDATE TBL_OrderSummary SET ErrorOnReview = 0 WHERE InvoiceID = ?"
        objCmd.Parameters.Append(objCmd.CreateParameter("invoiceid",3,1,15, var_invoiceid))
        objCmd.Execute()

        if add_reship_notes = "yes" then
            add_reship_notes = "Shipped replacement items (notated below) in invoice #" & move_to_invoice & "<br/>"
        else
            add_reship_notes = ""
        end if
        if add_storecredit_notes = "yes" then
            add_storecredit_notes = "Some items were out of stock and a " & var_refund_total & " store credit was issued<br/>"
        else
            add_storecredit_notes = ""
        end if
        if add_giftcert_notes = "yes" then
            add_giftcert_notes = "Some items were out of stock and a " & var_refund_total & " gift certificate was issued with code " & var_cert_code & "<br/>"
        else
            add_giftcert_notes = ""
        end if

        ' Notes for original order
        set objCmd = Server.CreateObject("ADODB.command")
        objCmd.ActiveConnection = DataConn
        objCmd.CommandText = "INSERT INTO tbl_invoice_notes (user_id, invoice_id, note) VALUES (?,?,?)"
        objCmd.Parameters.Append(objCmd.CreateParameter("user_id",3,1,10,user_id))
        objCmd.Parameters.Append(objCmd.CreateParameter("invoiceid",3,1,15,var_invoiceid))
        objCmd.Parameters.Append(objCmd.CreateParameter("note",200,1,1500, "Automated note: " & add_reship_notes & add_storecredit_notes & add_giftcert_notes))
        objCmd.Execute()

        ' =============== Notes for replacement/new order =========================
        set objCmd = Server.CreateObject("ADODB.command")
        objCmd.ActiveConnection = DataConn
        objCmd.CommandText = "INSERT INTO tbl_invoice_notes (user_id, invoice_id, note) VALUES (?,?,?)"
        objCmd.Parameters.Append(objCmd.CreateParameter("user_id",3,1,10,user_id))
        objCmd.Parameters.Append(objCmd.CreateParameter("invoiceid",3,1,15,move_to_invoice))
        objCmd.Parameters.Append(objCmd.CreateParameter("note",200,1,1500, "Automated note:  Replacement order for invoice #" & var_invoiceid))
        objCmd.Execute()

        ' ======== Encrypt data and write refund information to database so that customer can do a self serve refund =================

        if var_refund_total > 0 then
        
            Set objCrypt = Server.CreateObject("Bodyartforms.BAFCrypt")
            password = "3uBRUbrat77V"
            data = var_invoiceid & "|" & var_refund_total
            encrypted_code = objCrypt.Encrypt(password, data)

            set objCmd = Server.CreateObject("ADODB.command")
            objCmd.ActiveConnection = DataConn
            objCmd.CommandText = "INSERT INTO tbl_redeemable_refunds (invoice_id, refund_total, original_refund_total, gift_cert_id, encrypted_code, date_added) VALUES (?,?,?,?,?,?)"
            objCmd.Parameters.Append(objCmd.CreateParameter("invoice_id",3,1,15, var_invoiceid))
            objCmd.Parameters.Append(objCmd.CreateParameter("refund_total",6,1,20, var_refund_total))
            objCmd.Parameters.Append(objCmd.CreateParameter("original_refund_total",6,1,20, var_refund_total))
            objCmd.Parameters.Append(objCmd.CreateParameter("gift_cert_id",3,1,15, var_cert_id))
            objCmd.Parameters.Append(objCmd.CreateParameter("encrypted_code",200,1,250, encrypted_code))
            objCmd.Parameters.Append(objCmd.CreateParameter("date_added",200,1,75, now()))
            objCmd.Execute()

            Set objCrypt = Nothing

            ' decrypt refund information
            Set objCrypt = Server.CreateObject("Bodyartforms.BAFCrypt")
            password = "3uBRUbrat77V"
            data = encrypted_code
            decrypted = objCrypt.Decrypt(password, data)

            response.write "decrypted: " & decrypted
            Set objCrypt = Nothing

        end if '========== var_refund_total > 0

        mailer_type = "reship_approve"
        %>
        <!--#include virtual="emails/email_variables.asp"-->
        <%

    else '===== if agenda = deny

    end if '===== agenda approved/deny

end if '=== NOT rsGetInvoice.EOF


DataConn.Close()
%>