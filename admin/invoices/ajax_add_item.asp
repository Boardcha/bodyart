<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" %>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<!--#include virtual="/Connections/taxjar.asp"-->
<!--#include virtual="/taxjar/taxjar-nexus-values.asp"-->
<%
' SET VARIABLES
	amount_to_collect = 0
	state_tax_collectable = 0
	county_tax_collectable = 0 
	city_tax_collectable = 0
	special_district_tax_collectable = 0


	set objCmd = Server.CreateObject("ADODB.Command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "INSERT INTO TBL_OrderSummary(InvoiceID, ProductID, DetailID, qty, item_price) VALUES (?,?,?,?,?)"
	objCmd.Parameters.Append(objCmd.CreateParameter("invoiceid",3,1,15,request("invoiceid")))
	objCmd.Parameters.Append(objCmd.CreateParameter("productid",3,1,15,request("add_productid")))
	objCmd.Parameters.Append(objCmd.CreateParameter("add_detailid",3,1,15,request("add_detailid")))
	objCmd.Parameters.Append(objCmd.CreateParameter("add_qty",3,1,15,request("add_qty")))
	objCmd.Parameters.Append(objCmd.CreateParameter("add_price",6,1,10,request("add_price")))	
	objCmd.Execute()

	' Retrieve order
	Set objCmd = Server.CreateObject ("ADODB.Command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT qty FROM ProductDetails WHERE ProductDetailID = ?" 
	objCmd.Parameters.Append(objCmd.CreateParameter("detail_id",3,1,20, request("add_detailid") ))
	Set rsGetCurrentQty = objCmd.Execute()

	' Deduct inventory
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "UPDATE ProductDetails SET qty = qty - " & request("add_qty") & ", DateLastPurchased = '"& date() &"' WHERE ProductDetailID = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("detailid",3,1,15,request("add_detailid")))
	objCmd.Execute()

	'Write info to edits log	
	set objCmd = Server.CreateObject("ADODB.Command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "INSERT INTO tbl_edits_log (user_id, detail_id, description, edit_date) VALUES (" & user_id & ", " & request("add_detailid") & ",'Automated - Deducted " & request("add_qty") & " from stock. Updated from " & rsGetCurrentQty("qty") & " to " & rsGetCurrentQty("qty") - request("add_qty") & " - adding item to order from invoice edit page','" & now() & "')"
	objCmd.Execute()
	Set objCmd = Nothing
	
	' Retrieve order
	Set objCmd = Server.CreateObject ("ADODB.Command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT * FROM sent_items WHERE ID = ?" 
	objCmd.Parameters.Append(objCmd.CreateParameter("string_id",3,1,12,request("invoiceid")))
	Set rsGetOrder = objCmd.Execute()
	
	set objCmd = Server.CreateObject("ADODB.Command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT DiscountPercent FROM TBLDiscounts WHERE DiscountCode = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("coupon_code",200,1,50,rsGetOrder.Fields.Item("coupon_code").Value))
	Set rsGetCouponDiscount = objCmd.Execute()
	
	item_price = request("add_price")
	' If there was a coupon used then add the coupon and tax amount on the order
	if rsGetOrder.Fields.Item("coupon_code").Value <> "" then
	
		if NOT rsGetCouponDiscount.eof then
			item_price = FormatNumber((request("add_price") - ((rsGetCouponDiscount.Fields.Item("DiscountPercent").Value / 100) * request("add_price"))) * request("add_qty"), -1, -2, -0, -2)
			
			var_discount_difference = request("add_price") * request("add_qty") - item_price
			discount = ""
			
			if rsGetOrder.Fields.Item("coupon_code").Value = "YTG89R57" then
				discount = "total_preferred_discount = total_preferred_discount + " & var_discount_difference
			else
				discount = "total_coupon_discount = total_coupon_discount + " & var_discount_difference
			end if
		else
			item_price = request("add_price") * request("add_qty")
		end if
	
	
'	response.write "form price: " & request("add_price") & "<br/>"
'	response.write "item_price: " & item_price & "<br/>"	
'	response.write "discount: " & discount & "<br/>"
'	response.write "var_discount_difference: " & var_discount_difference & "<br/>"


		set objCmd = Server.CreateObject("ADODB.Command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "UPDATE sent_items SET " &  discount & " WHERE ID = ?" 
		objCmd.Parameters.Append(objCmd.CreateParameter("string_id",3,1,12,request("invoiceid")))
		objCmd.Execute()

	end if	' if coupon needs to be calculated

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
			"""amount"":""" & item_price * request("add_qty") & """," & _
			"""line_items"": [{" & _
				"""id"":""1""," & _
				"""quantity"": " & request("add_qty") & "," & _
				"""unit_price"": " & item_price & "," & _
				"""discount"": 0" & _
			"}]," & _
			"""nexus_addresses"": [" & _
				taxjar_nexus_values & _
			"]" & _
			"}")

		response.write HttpReq.responseText

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

			response.write "X value - " & x & ",<br>"

				if instr(x,"amount_to_collect") > 0 then
					amount_to_collect = Split(x, ":")(1)
				end if
				if instr(x,"state_tax_collectable") > 0 then
					state_tax_collectable = Split(x, ":")(1)
				end if
				if instr(x,"county_tax_collectable") > 0  then
					county_tax_collectable = Split(x, ":")(1)
				end if
				if instr(x,"city_tax_collectable") > 0  then
					city_tax_collectable = Split(x, ":")(1)
				end if
				if instr(x,"special_district_tax_collectable") > 0 then
					special_district_tax_collectable = Split(x, ":")(1)
				end if
		next
		set HttpReq = Nothing

		'response.write "<br>COUNTRY - " & taxjar_to_country
		'response.write "<BR>TAXJAR OUTPUT amount_to_collect - " & amount_to_collect
		'response.write "<BR>TAXJAR OUTPUT state_tax_collectable - " & state_tax_collectable
		'response.write "<BR>TAXJAR OUTPUT county_tax_collectable - " & county_tax_collectable
		'response.write "<BR>TAXJAR OUTPUT city_tax_collectable - " & city_tax_collectable
		'response.write "<BR>TAXJAR OUTPUT special_district_tax_collectable - " & special_district_tax_collectable


	set objCmd = Server.CreateObject("ADODB.Command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "UPDATE sent_items SET total_sales_tax = total_sales_tax + " & amount_to_collect & ", taxes_state_only = taxes_state_only + " & state_tax_collectable & ", taxes_county_only = taxes_county_only + " & county_tax_collectable & ", taxes_city_only = taxes_city_only + " & city_tax_collectable & ", taxes_special_only = taxes_special_only + " & special_district_tax_collectable & " WHERE ID = ?" 
	objCmd.Parameters.Append(objCmd.CreateParameter("string_id",3,1,12,request("invoiceid")))
	objCmd.Execute()

end if ' if there are taxes to calc
				
DataConn.Close()
%>