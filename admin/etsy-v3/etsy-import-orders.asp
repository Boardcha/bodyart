<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/admin/etsy-v3/etsy-refresh-token.asp" -->
<%
set rest = Server.CreateObject("Chilkat_9_5_0.Rest")

autoReconnect = 1
tls = 1
success = rest.Connect("openapi.etsy.com",443,tls,autoReconnect)
If (success = 0) Then
    Response.Write "<pre>" & Server.HTMLEncode( rest.LastErrorText) & "</pre>"
    Response.End
End If

set sbAuthHeaderVal = Server.CreateObject("Chilkat_9_5_0.StringBuilder")
success = sbAuthHeaderVal.Append("Bearer ")
success = sbAuthHeaderVal.Append(etsy_access_token)
rest.Authorization = sbAuthHeaderVal.GetAsString() 

success = rest.AddQueryParam("client_id", etsy_consumer_key) 
success = rest.AddQueryParam("was_shipped",0)
success = rest.AddQueryParam("was_paid",1)
success = rest.AddQueryParam("limit",100)

jsonResponseText = rest.FullRequestNoBody("GET","/v3/application/shops/" & etsy_baf_shop_id & "/receipts")
If (rest.LastMethodSuccess = 0) Then
    Response.Write "<pre>" & Server.HTMLEncode( rest.LastErrorText) & "</pre>"
    Response.End
End If 
        
set jsonResponse = Server.CreateObject("Chilkat_9_5_0.JsonObject")
success = jsonResponse.Load(jsonResponseText)
jsonResponse.EmitCompact = 0


'Response.Write "<pre>" & Server.HTMLEncode( jsonResponse.Emit()) & "</pre>"
'Response.Write "<pre>" & Server.HTMLEncode( "Response status code: " & rest.ResponseStatusCode) & "</pre>"

i = 0
count_i = jsonResponse.SizeOfArray("results")
Do While i < count_i
	jsonResponse.I = i
    var_email = jsonResponse.StringOf("results[i].buyer_email") 
    split_name = split(jsonResponse.StringOf("results[i].name"), " ")
        var_first = split_name(0)
        if uBound(split_name) > 0 then
        var_last = split_name(1)
        end if
    var_address1 = jsonResponse.StringOf("results[i].first_line") 
    var_address2 = replace(jsonResponse.StringOf("results[i].second_line"), "null", "")
    var_city = jsonResponse.StringOf("results[i].city") 
    var_state = jsonResponse.StringOf("results[i].state") 
    var_zip = jsonResponse.StringOf("results[i].zip") 
	var_country = jsonResponse.StringOf("results[i].country_iso") 
    var_receipt_id = jsonResponse.StringOf("results[i].receipt_id")
    var_shipping_rate = jsonResponse.IntOf("results[i].total_shipping_cost.amount") 
	If jsonResponse.IntOf("results[i].total_shipping_cost.divisor") > 0 Then var_shipping_rate = jsonResponse.IntOf("results[i].total_shipping_cost.amount") / jsonResponse.IntOf("results[i].total_shipping_cost.divisor")	
    var_order_tax = jsonResponse.IntOf("results[i].total_tax_cost.amount") 
	If jsonResponse.IntOf("results[i].total_tax_cost.divisor") > 0 Then var_order_tax = jsonResponse.IntOf("results[i].total_tax_cost.amount") / jsonResponse.IntOf("results[i].total_tax_cost.divisor")
		
	
    if var_shipping_rate = 3.95 then

    var_shipping_type = "DHL Basic mail"
    
    elseif var_shipping_rate = 0.00 then
    
    var_shipping_type = "DHL Basic mail"
    
    elseif var_shipping_rate = 4.95 then
    
    var_shipping_type = "DHL Expedited Max"
    
    elseif var_shipping_rate = 5.95 then
    
    var_shipping_type = "USPS First Class Mail"
    
    elseif var_shipping_rate = 7.95 then
    
    var_shipping_type = "USPS Priority mail"
    
    elseif var_shipping_rate = 23.95 then
    
    var_shipping_type = "USPS Express mail"
    
    elseif var_shipping_rate = 8.95 then
    
    var_shipping_type = "DHL GlobalMail Parcel Priority"
    
    end if

	
    ' Only insert record if there is no transaction ID already in the table 
    set objCmd = Server.CreateObject("ADODB.Command")
    objCmd.ActiveConnection = DataConn
    objCmd.CommandText = "SELECT transactionid FROM sent_items where transactionid = ?"
    objCmd.Parameters.Append(objCmd.CreateParameter("var_receipt_id",200,1,100,var_receipt_id))
    set rsCheckDupeOrder = objCmd.Execute()

    if rsCheckDupeOrder.eof then

    set objCmd = Server.CreateObject("ADODB.Command")
    objCmd.ActiveConnection = DataConn
    objCmd.CommandText = "INSERT INTO sent_items (shipped, pay_method, ship_code, email, customer_first, customer_last, address, address2, city, state, zip, country, transactionID, date_order_placed, shipping_rate, shipping_type, total_sales_tax) VALUES ('Pending shipment', 'Etsy', 'paid', ?,?,?,?,?,?,?,?,?,?,'" & now() & "' ,?,?,?)"
    objCmd.Parameters.Append(objCmd.CreateParameter("email",200,1,100,var_email))
    objCmd.Parameters.Append(objCmd.CreateParameter("first",200,1,50, replace(var_first,"""", "")))
    objCmd.Parameters.Append(objCmd.CreateParameter("last",200,1,50, replace(var_last,"""", "")))
    objCmd.Parameters.Append(objCmd.CreateParameter("address",200,1,100,var_address1))
    objCmd.Parameters.Append(objCmd.CreateParameter("address2",200,1,100,var_address2))
    objCmd.Parameters.Append(objCmd.CreateParameter("city",200,1,100,var_city))
    objCmd.Parameters.Append(objCmd.CreateParameter("state",200,1,50,var_state))
    objCmd.Parameters.Append(objCmd.CreateParameter("zip",200,1,15,var_zip))
    objCmd.Parameters.Append(objCmd.CreateParameter("country",200,1,5,var_country))
    objCmd.Parameters.Append(objCmd.CreateParameter("var_receipt_id",200,1,100,var_receipt_id))
    objCmd.Parameters.Append(objCmd.CreateParameter("shipping_rate",6,1,10,var_shipping_rate))
    objCmd.Parameters.Append(objCmd.CreateParameter("shipping_type",200,1,100,var_shipping_type))
    objCmd.Parameters.Append(objCmd.CreateParameter("total_sales_tax",6,1,10,var_order_tax))
    objCmd.Execute()
    Set objCmd = Nothing


	
    '-------- Get invoice # for items ---------------
    set objCmd = Server.CreateObject("ADODB.Command")
    objCmd.ActiveConnection = DataConn
    objCmd.CommandText = "SELECT id FROM sent_items WHERE transactionID = ? ORDER BY ID DESC"
    objCmd.Parameters.Append(objCmd.CreateParameter("var_receipt_id",200,1,100,var_receipt_id))
    set rsGetInvoiceNum = objCmd.Execute()
        if NOT rsGetInvoiceNum.eof then
            var_invoicenum = rsGetInvoiceNum.Fields.Item("id").Value
        else
            var_invoicenum = 0
        end if   


    success = rest.ClearAllQueryParams()
	success = rest.AddQueryParam("client_id", etsy_consumer_key) 
    
    jsonItemsResponseText = rest.FullRequestNoBody("GET","/v3/application/shops/" & etsy_baf_shop_id & "/receipts/" & var_receipt_id & "/transactions")
    If (rest.LastMethodSuccess = 0) Then
        Response.Write "<pre>" & Server.HTMLEncode( rest.LastErrorText) & "</pre>"
        Response.End
    End If
    
            
    set jsonItemResponse = Server.CreateObject("Chilkat_9_5_0.JsonObject")
    success = jsonItemResponse.Load(jsonItemsResponseText)
    jsonItemResponse.EmitCompact = 0
    
    'Response.Write "<pre>" & Server.HTMLEncode( jsonItemResponse.Emit()) & "</pre>"
    'Response.Write "<pre>" & Server.HTMLEncode( "Response status code: " & rest.ResponseStatusCode) & "</pre>"

    j = 0
    count_j = jsonItemResponse.SizeOfArray("results")
    Do While j < count_j
        jsonItemResponse.J = j


		'var_product_detailid = jsonItemResponse.IntOf("results[j].product_data.sku")
		'There were a sku element in transactions in the API v2 (above line) which is not exist in V3. 
		'Etsy developers says they will include this again, so we may not have to send nested requests to get sku when they applied this.
		'https://github.com/etsy/open-api/issues/161
		' === Start getting sku value of the variant
		success = rest.ClearAllQueryParams()
		success = rest.AddQueryParam("client_id", etsy_consumer_key) 
		etsy_listing_id = jsonItemResponse.IntOf("results[j].listing_id")
		etsy_product_id = jsonItemResponse.StringOf("results[j].product_id")
		jsonSkuResponseText = rest.FullRequestNoBody("GET","/v3/application/listings/" & etsy_listing_id & "/inventory/products/" & etsy_product_id)
		If (rest.LastMethodSuccess = 0) Then
			Response.Write "<pre>" & Server.HTMLEncode( rest.LastErrorText) & "</pre>"
			Response.End
		End If
		set jsonSkuResponse = Server.CreateObject("Chilkat_9_5_0.JsonObject")
		success = jsonSkuResponse.Load(jsonSkuResponseText)
		jsonSkuResponse.EmitCompact = 0
		var_product_detailid = jsonSkuResponse.StringOf("sku")
		' === End of getting sku of the variant
		
		etsy_qty = jsonItemResponse.IntOf("results[j].quantity")
		var_item_title = jsonItemResponse.StringOf("results[j].title")
		var_etsy_price = jsonItemResponse.IntOf("results[j].price.amount")
		if jsonItemResponse.IntOf("results[j].price.divisor") > 0 Then var_etsy_price = jsonItemResponse.IntOf("results[j].price.amount") / jsonItemResponse.IntOf("results[j].price.divisor")
		var_productid = 0

		
		'============ Search Etsy title for character that tells us whether to deduct 1 or 2 from our site for Etsy items sold ===============
		if InStr(var_item_title, ":") > 0 then
			our_qty = etsy_qty
			var_item_price = var_etsy_price
		else
			our_qty = 2 * etsy_qty
			var_item_price = var_etsy_price / 2
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
     
    
        '------- Insert order items into table ---------------
        set objCmd = Server.CreateObject("ADODB.Command")
        objCmd.ActiveConnection = DataConn
        objCmd.CommandText = "INSERT INTO TBL_OrderSummary (InvoiceID, detail_transactionid, DetailID, ProductID, item_price, qty) VALUES (?,?,?,?,?,?)"
        objCmd.Parameters.Append(objCmd.CreateParameter("invoiceid",3,1,15, var_invoicenum))
        objCmd.Parameters.Append(objCmd.CreateParameter("var_receipt_id",200,1,100,var_receipt_id))
        objCmd.Parameters.Append(objCmd.CreateParameter("product_detailid",3,1,15,var_product_detailid))
        objCmd.Parameters.Append(objCmd.CreateParameter("productid",3,1,15, var_productid))
        objCmd.Parameters.Append(objCmd.CreateParameter("item_price",6,1,10,var_item_price))
        objCmd.Parameters.Append(objCmd.CreateParameter("qty",3,1,10,our_qty))
        objCmd.Execute()

        '------- Deduct quantities on order items ---------------
        set objCmd = Server.CreateObject("ADODB.Command")
        objCmd.ActiveConnection = DataConn
        objCmd.CommandText = "UPDATE ProductDetails SET qty = qty - " & our_qty & " WHERE ProductDetailID = ?"
        objCmd.Parameters.Append(objCmd.CreateParameter("product_detailid",200,1,100,var_product_detailid))
        objCmd.Execute()

        Set objCmd = Nothing
        j = j + 1
    Loop

    '===== CHECK TO SEE IF ANY PAYMENTS ARE IN THE SYSTEM FOR THE TRANSACTION. THIS IS REALLY THE ONLY WAY TO SEE IF AN ORDER HAS BEEN CANCELED AS OF APRIL 2020. THEY HAVE NO WAY TO SEARCH DIRECTLY FROM CANCELED ORDERS =======
    success = rest.ClearAllQueryParams()
	success = rest.AddQueryParam("client_id", etsy_consumer_key) 
    jsonCancelledResponseText = rest.FullRequestNoBody("GET","/v3/application/shops/" & etsy_baf_shop_id & "/receipts/" & var_receipt_id & "/payments")
    If (rest.LastMethodSuccess = 0) Then
        Response.Write "<pre>" & Server.HTMLEncode( rest.LastErrorText) & "</pre>"
        Response.End
    End If
    
            
    set jsonCanceledResponse = Server.CreateObject("Chilkat_9_5_0.JsonObject")
    success = jsonCanceledResponse.Load(jsonCancelledResponseText)
    jsonCanceledResponse.EmitCompact = 0
    
    'Response.Write "<pre>" & Server.HTMLEncode( jsonCanceledResponse.Emit()) & "</pre>"
    'Response.Write "<pre>" & Server.HTMLEncode( "Response status code: " & rest.ResponseStatusCode) & "</pre>"

    j = 0
    count_j = jsonCanceledResponse.SizeOfArray("results")
    Do While j < count_j
        jsonCanceledResponse.J = j

            payment_id = jsonCanceledResponse.StringOf("results[j].payment_id")

        Set objCmd = Nothing
        j = j + 1
    Loop
	
    '======= SET ORDER IN OUR ADMIN TO CANCELLED, NOT PAID STATUS ============
    if payment_id = "" then
        set objCmd = Server.CreateObject("ADODB.command")
        objCmd.ActiveConnection = DataConn
        objCmd.CommandText = "UPDATE sent_items SET shipped = 'Cancelled', ship_code = 'not paid' WHERE ID = ?"
        objCmd.Parameters.Append(objCmd.CreateParameter("invoice_id",3,1,15, var_invoicenum ))
        objCmd.Execute()
    end if 
    '============== END CANCEL ORDER CHECK ===============================================

    '------- Add gauge card to order  ---------------
    set objCmd = Server.CreateObject("ADODB.Command")
    objCmd.ActiveConnection = DataConn
    objCmd.CommandText = "INSERT INTO TBL_OrderSummary (InvoiceID, detail_transactionid, DetailID, ProductID, item_price, qty) VALUES (?,?,5461,1430,0,1)"
    objCmd.Parameters.Append(objCmd.CreateParameter("invoiceid",3,1,15, var_invoicenum))
    objCmd.Parameters.Append(objCmd.CreateParameter("var_receipt_id",200,1,100,var_receipt_id))
    objCmd.Execute()

    '------- Add random sticker to order  ---------------
    set objCmd = Server.CreateObject("ADODB.Command")
    objCmd.ActiveConnection = DataConn
    objCmd.CommandText = "INSERT INTO TBL_OrderSummary (InvoiceID, detail_transactionid, DetailID, ProductID, item_price, qty) VALUES (?,?,72198,3928,0,1)"
    objCmd.Parameters.Append(objCmd.CreateParameter("invoiceid",3,1,15, var_invoicenum))
    objCmd.Parameters.Append(objCmd.CreateParameter("var_receipt_id",200,1,100,var_receipt_id))
    objCmd.Execute()

    end if ' do not insert duplicate order
    set rsCheckDupeOrder = nothing

    i = i + 1

Loop

Set rest = Nothing
Set sbAuthHeaderVal = Nothing
Set jsonResponse = Nothing
Set rsCheckDupeOrder = Nothing
Set rsGetInvoiceNum = Nothing
Set jsonItemResponse = Nothing
Set jsonSkuResponse = Nothing
Set jsonCanceledResponse = Nothing
Set objCmd = Nothing
DataConn.Close

%>