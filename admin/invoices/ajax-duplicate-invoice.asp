<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" %>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
var_orig_invoiceid = request.form("invoiceid")
    '==== COPY INVOICE =======
    set objCmd = Server.CreateObject("ADODB.Command")
    objCmd.ActiveConnection = MM_bodyartforms_sql_STRING
    objCmd.CommandText = "INSERT INTO sent_items (shipped, customer_ID, customer_first, customer_last, company, address, address2, city, state, province, zip, country, email, date_order_placed, shipping_rate, ship_code, phone, transactionID, pay_method, shipping_type, autoclave, anodize) SELECT 'Pending...', customer_ID, customer_first, customer_last, company, address, address2, city, state, province, zip, country, email, '" & now() & "', 0, 'paid', phone, transactionID, pay_method, shipping_type, autoclave, anodize FROM sent_items WHERE ID = ?"
    objCmd.Parameters.Append(objCmd.CreateParameter("invoiceid",3,1,15, var_orig_invoiceid ))
    objCmd.Execute() 

    '==== GET THE NEW INVOICE # TO COPY THE ITEMS INTO IT ======
    Set objCmd = Server.CreateObject ("ADODB.Command")
    objCmd.ActiveConnection = MM_bodyartforms_sql_STRING
    objCmd.CommandText = "SELECT ID, email FROM dbo.sent_items WHERE email = ? ORDER BY ID DESC" 
    objCmd.Parameters.Append(objCmd.CreateParameter("email",200,1,250, request.form("email") ))
    Set rsGetNewInvoiceID = objCmd.Execute

    '===== Copy ordered items ======
    set objCmd = Server.CreateObject("ADODB.Command")
    objCmd.ActiveConnection = MM_bodyartforms_sql_STRING
    objCmd.CommandText = "INSERT INTO TBL_OrderSummary (InvoiceID, ProductID, DetailID, qty, item_price, PreOrder_Desc, anodization_id_ordered, anodization_fee) SELECT " & rsGetNewInvoiceID("ID") & ", ProductID, DetailID, qty, 0, PreOrder_Desc, anodization_id_ordered, anodization_fee FROM TBL_OrderSummary WHERE InvoiceID = ?" 
    objCmd.Parameters.Append(objCmd.CreateParameter("invoiceid",3,1,15, var_orig_invoiceid ))
    objCmd.Execute() 
    var_new_invoiceid = rsGetNewInvoiceID("ID")

    '===== RETRIEVE ITEM DETAILS FROM ORIGINAL INVOICE TO DEDUCT QUANTITY ON DUPLICATE =========
    set objCmd = Server.CreateObject("ADODB.Command")
    objCmd.ActiveConnection = MM_bodyartforms_sql_STRING
    objCmd.CommandText = "SELECT DetailID, qty FROM TBL_OrderSummary WHERE InvoiceID = ?"
    objCmd.Parameters.Append(objCmd.CreateParameter("invoiceid",3,1,15, var_orig_invoiceid ))
    Set rsGetOrderDetails = objCmd.Execute()  

    While NOT rsGetOrderDetails.EOF

        '===== DEDUCT INVENTORY =============
        set objCmd = Server.CreateObject("ADODB.command")
        objCmd.ActiveConnection = DataConn
        objCmd.CommandText = "UPDATE ProductDetails SET qty = qty - " & rsGetOrderDetails("qty") & ", DateLastPurchased = '" & now() & "' WHERE ProductDetailID = " & rsGetOrderDetails("DetailID")
        objCmd.Execute()

        '========  Write info to edits log	 ==========
        set objCmd = Server.CreateObject("ADODB.Command")
        objCmd.ActiveConnection = DataConn
        objCmd.CommandText = "INSERT INTO tbl_edits_log (user_id, detail_id, description, edit_date) VALUES (" & user_id & ", " & rsGetOrderDetails("DetailID") & ",'Automated - Deducted " & rsGetOrderDetails("qty") & " from stock -- duplicating invoice','" & now() & "')"
        objCmd.Execute()
        Set objCmd = Nothing

    rsGetOrderDetails.MoveNext()
    Wend

    '===== AUTOMATED NOTES LOG FOR OLD ORDER =================
    set objCmd = Server.CreateObject("ADODB.command")
    objCmd.ActiveConnection = DataConn
    objCmd.CommandText = "INSERT INTO tbl_invoice_notes (user_id, invoice_id, note) VALUES (?,?, 'Automated message (duplicate invoice button): Created new duplicated order " & var_new_invoiceid & "')"
    objCmd.Parameters.Append(objCmd.CreateParameter("user_id",3,1,10, user_id))
    objCmd.Parameters.Append(objCmd.CreateParameter("invoiceid",3,1,15, var_orig_invoiceid))
    objCmd.Execute()

    '===== AUTOMATED NOTES LOG FOR NEW ORDER =================
    set objCmd = Server.CreateObject("ADODB.command")
    objCmd.ActiveConnection = DataConn
    objCmd.CommandText = "INSERT INTO tbl_invoice_notes (user_id, invoice_id, note) VALUES (?,?, 'Automated message (duplicate invoice button): Created new duplicated order from invoice " & var_orig_invoiceid & ". Quantities have been automatically deducted.')"
    objCmd.Parameters.Append(objCmd.CreateParameter("user_id",3,1,10, user_id))
    objCmd.Parameters.Append(objCmd.CreateParameter("invoiceid",3,1,15, var_new_invoiceid))
    objCmd.Execute()
%>
{
    "new_invoiceid" : "<%= var_new_invoiceid %>"
}
<%
DataConn.Close()
%>