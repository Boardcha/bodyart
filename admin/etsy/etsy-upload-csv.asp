<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" %>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
' =========================================================================================
' Upload .csv file for Etsy orders and order details

' Link below to install Microsoft Access Database Engine 2016 Redistributable
' https://www.microsoft.com/en-us/download/details.aspx?id=54920
' =========================================================================================

%>
{
<%
Dim strConn, conn, rs

' --------------------------------------------
' Upload CSV file
' --------------------------------------------
Set Upload = Server.CreateObject("Persits.Upload")
	Upload.OverwriteFiles = True
    'Upload.Save("C:\inetpub\wwwroot\bootstrap-svn\admin\etsy") 'LOCALHOST TESTING
    Upload.Save("C:\inetpub\bootstrap-baf\admin\etsy")  'LIVE SERVER


    strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & _
    Server.MapPath("\admin\etsy\") & ";Extended Properties=""text;HDR=Yes;FMT-Delimited"";"
    Set conn = Server.CreateObject("ADODB.Connection")
    conn.Open strConn

    x = 1
    For Each File in Upload.Files
        'File.Copy "C:\inetpub\wwwroot\bootstrap-svn\admin\etsy\etsy" & x & ".csv"   'LOCALHOST
        File.Copy "C:\inetpub\bootstrap-baf\admin\etsy\etsy" & x & ".csv"  'LIVE SERVER
        File.Delete        
    
        ' -----------------------------------------------------------------------
        ' Connect to CSV file to find out if file is for orders, or order details
        ' -----------------------------------------------------------------------     

        Set rsDetectFile = Server.CreateObject("ADODB.recordset")
        rsDetectFile.open "SELECT * FROM etsy" & x & ".csv", conn

        For Each header In rsDetectFile.Fields
            If header.Name = "Count - Number of Items" Then
                order_file = "etsy" & x
                'Response.Write("orders file: " & order_file)
                Exit For
            End If
        Next

        For Each header In rsDetectFile.Fields
            If header.Name = "Item - SKU" Then
                detail_file = "etsy" & x
                'Response.Write("details file: " & detail_file)
                Exit For
            End If
        Next

        x = x + 1
    next ' for each file uploaded
    
    ' -----------------------------------------------------------------------
    ' Insert orders into database from order csv file
    ' -----------------------------------------------------------------------

    Set rsGetOrders = Server.CreateObject("ADODB.recordset")
    rsGetOrders.open "SELECT * FROM " & order_file & ".csv", conn 
    var_record = 1
    while not rsGetOrders.eof
        'response.write "record inserted " & var_record & "<br>"
        var_email = rsGetOrders.Fields.Item("Customer Email").Value 
        
        split_name = split(rsGetOrders.Fields.Item("Ship To - Name").Value, " ")
            var_first = split_name(0)
            if uBound(split_name) > 0 then
            var_last = split_name(1)
            end if
        var_address1 = rsGetOrders.Fields.Item("Ship To - Address 1").Value
        var_address2 = rsGetOrders.Fields.Item("Ship To - Address 2").Value
        var_city = rsGetOrders.Fields.Item("Ship To - City").Value
        var_state = rsGetOrders.Fields.Item("Ship To - State").Value
        var_zip = rsGetOrders.Fields.Item("Ship To - Postal Code").Value
        var_country = rsGetOrders.Fields.Item("Ship To - Country").Value
        var_phone = rsGetOrders.Fields.Item("Ship To - Phone").Value
        var_date_order_placed = rsGetOrders.Fields.Item("Date - Order Date").Value
        var_transactionid = rsGetOrders.Fields.Item("Order - Number").Value
        var_shipping_rate = rsGetOrders.Fields.Item("Amount - Order Shipping").Value
        var_order_tax = rsGetOrders.Fields.Item("Amount - Order Tax").Value

        if var_shipping_rate = 4.95 then

        var_shipping_type = "DHL Basic mail"
        
        elseif var_shipping_rate = 0.00 then
        
        var_shipping_type = "DHL Basic mail"
        
        elseif var_shipping_rate = 5.95 then
        
        var_shipping_type = "DHL Expedited Max"
        
        elseif var_shipping_rate = 7.95 then
        
        var_shipping_type = "USPS Priority mail"
        
        elseif var_shipping_rate = 23.95 then
        
        var_shipping_type = "USPS Express mail"
        
        elseif var_shipping_rate = 7.99 then
        
        var_shipping_type = "DHL GlobalMail Parcel Priority"
        
        end if

        ' Only insert record if there is no transaction ID already in the table 
        set objCmd = Server.CreateObject("ADODB.Command")
        objCmd.ActiveConnection = DataConn
        objCmd.CommandText = "SELECT transactionid FROM sent_items where transactionid = ?"
        objCmd.Parameters.Append(objCmd.CreateParameter("transactionid",200,1,100,var_transactionid))
        set rsCheckDupeOrder = objCmd.Execute()

        if rsCheckDupeOrder.eof then

        set objCmd = Server.CreateObject("ADODB.Command")
        objCmd.ActiveConnection = DataConn
        objCmd.CommandText = "INSERT INTO sent_items (shipped, pay_method, ship_code, email, customer_first, customer_last, address, address2, city, state, zip, country, phone, transactionID, date_order_placed, shipping_rate, shipping_type, total_sales_tax) VALUES ('Pending shipment', 'Etsy', 'paid', ?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"
        objCmd.Parameters.Append(objCmd.CreateParameter("email",200,1,100,var_email))
        objCmd.Parameters.Append(objCmd.CreateParameter("first",200,1,50, replace(var_first,"""", "")))
        objCmd.Parameters.Append(objCmd.CreateParameter("last",200,1,50, replace(var_last,"""", "")))
        objCmd.Parameters.Append(objCmd.CreateParameter("address",200,1,100,var_address1))
        objCmd.Parameters.Append(objCmd.CreateParameter("address2",200,1,100,var_address2))
        objCmd.Parameters.Append(objCmd.CreateParameter("city",200,1,100,var_city))
        objCmd.Parameters.Append(objCmd.CreateParameter("state",200,1,50,var_state))
        objCmd.Parameters.Append(objCmd.CreateParameter("zip",200,1,15,var_zip))
        objCmd.Parameters.Append(objCmd.CreateParameter("country",200,1,5,var_country))
        objCmd.Parameters.Append(objCmd.CreateParameter("phone",200,1,30,var_phone))
        objCmd.Parameters.Append(objCmd.CreateParameter("transactionid",200,1,100,var_transactionid))
        objCmd.Parameters.Append(objCmd.CreateParameter("date_placed",200,1,100,var_date_order_placed))
        objCmd.Parameters.Append(objCmd.CreateParameter("shipping_rate",6,1,10,var_shipping_rate))
        objCmd.Parameters.Append(objCmd.CreateParameter("shipping_type",200,1,100,var_shipping_type))
        objCmd.Parameters.Append(objCmd.CreateParameter("total_sales_tax",6,1,10,var_order_tax))
        objCmd.Execute()
        Set objCmd = Nothing

        end if ' do not insert duplicate order
        var_record = var_record + 1
    rsGetOrders.movenext
    wend
    set rsCheckDupeOrder = nothing

    ' -----------------------------------------------------------------------
    ' Insert orders into database from order DETAILS csv file
    ' -----------------------------------------------------------------------

    Set rsGetDetails = Server.CreateObject("ADODB.recordset")
    rsGetDetails.open "SELECT * FROM " & detail_file & ".csv", conn 
    
    var_details = 1
    while not rsGetDetails.eof
    'response.write "Detail written to db " & var_details & "<br>"
        var_detail_transactionid = rsGetDetails.Fields.Item("Order - Number").Value
        var_product_detailid = rsGetDetails.Fields.Item("Item - SKU").Value
        etsy_qty = rsGetDetails.Fields.Item("Item - Qty").Value
        var_productid = 0

        '============ Search Etsy title for character that tells us whether to deduct 1 or 2 from our site for Etsy items sold ===============
        if InStr(rsGetDetails.Fields.Item("Item - Name").Value, ":") > 0 then
            our_qty = etsy_qty
            var_item_price = rsGetDetails.Fields.Item("Item - Price").Value
        else
            our_qty = 2 * etsy_qty
            var_item_price = rsGetDetails.Fields.Item("Item - Price").Value / 2
        end if


            ' Get productid to insert into table
            set objCmd = Server.CreateObject("ADODB.Command")
            objCmd.ActiveConnection = DataConn
            objCmd.CommandText = "SELECT ProductID FROM ProductDetails WHERE ProductDetailID = ?"
            objCmd.Parameters.Append(objCmd.CreateParameter("detailid",3,1,15,var_product_detailid))
            set rsProductID = objCmd.Execute()

            if NOT rsProductID.eof then
                var_productid = rsProductID.Fields.Item("ProductID").Value
            else
                var_productid = 0
            end if

            '-------- Get invoice # for items ---------------
            set objCmd = Server.CreateObject("ADODB.Command")
            objCmd.ActiveConnection = DataConn
            objCmd.CommandText = "SELECT id FROM sent_items WHERE transactionID = ?"
            objCmd.Parameters.Append(objCmd.CreateParameter("detail_transactionid",200,1,100,var_detail_transactionid))
            set rsGetInvoiceNum = objCmd.Execute()
                if NOT rsGetInvoiceNum.eof then
                    var_invoicenum = rsGetInvoiceNum.Fields.Item("id").Value
                else
                    var_invoicenum = 0
                end if          
            
        
            '------- Insert order items into table ---------------
            set objCmd = Server.CreateObject("ADODB.Command")
            objCmd.ActiveConnection = DataConn
            objCmd.CommandText = "INSERT INTO TBL_OrderSummary (InvoiceID, detail_transactionid, DetailID, ProductID, item_price, qty) VALUES (?,?,?,?,?,?)"
            objCmd.Parameters.Append(objCmd.CreateParameter("invoiceid",3,1,15, var_invoicenum))
            objCmd.Parameters.Append(objCmd.CreateParameter("detail_transactionid",200,1,100,var_detail_transactionid))
            objCmd.Parameters.Append(objCmd.CreateParameter("product_detailid",3,1,15,var_product_detailid))
            objCmd.Parameters.Append(objCmd.CreateParameter("productid",3,1,15, var_productid))
            objCmd.Parameters.Append(objCmd.CreateParameter("item_price",6,1,10,var_item_price))
            objCmd.Parameters.Append(objCmd.CreateParameter("qty",3,1,10,our_qty))
            objCmd.Execute()

            ' -----------------------------------------------------------------------
            ' Recode detail invoice ID's with the correct # from the orders table
            ' ----------------------------------------------------------------------
            'set objCmd = Server.CreateObject("ADODB.Command")
            'objCmd.ActiveConnection = DataConn
            'objCmd.CommandText = "UPDATE TBL_OrderSummary SET InvoiceID = i.ID FROM sent_items 'as i INNER JOIN TBL_OrderSummary as d ON i.ID = d.InvoiceID WHERE ? = 'i.transactionid AND i.pay_method = 'Etsy'"
            'objCmd.Parameters.Append(objCmd.CreateParameter("detail_transactionid",200,1,100,'var_detail_transactionid))
            'objCmd.Execute()
            'Set objCmd = Nothing

            '------- Deduct quantities on order items ---------------
            set objCmd = Server.CreateObject("ADODB.Command")
            objCmd.ActiveConnection = DataConn
            objCmd.CommandText = "UPDATE ProductDetails SET qty = qty - " & our_qty & " WHERE ProductDetailID = ?"
            objCmd.Parameters.Append(objCmd.CreateParameter("product_detailid",200,1,100,var_product_detailid))
            objCmd.Execute()

            Set objCmd = Nothing
            set rsProductID = nothing
            set rsGetInvoiceNum = nothing
            var_details = var_details + 1
    rsGetDetails.movenext
    wend

    set rsCheckDupeDetail = nothing

%>
    "status":"success",
    "reason":""
}
<%

conn.close()
DataConn.Close()
%>