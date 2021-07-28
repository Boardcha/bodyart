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

	
	' Get information to deduct inventory
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT qty, DetailID FROM QRY_OrderDetails WHERE OrderDetailID = ?" 
	objCmd.Parameters.Append(objCmd.CreateParameter("id",3,1,10,request("detailid")))
	set rsUpdate = objCmd.Execute()
	
	' Put item back into stock
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "UPDATE ProductDetails SET qty = qty + " & rsUpdate.Fields.Item("qty").Value & ", active = 1 WHERE ProductDetailID = " & rsUpdate.Fields.Item("DetailID").Value
	objCmd.Execute()
	
	' Reactive main product listing
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn  
	objCmd.CommandText = "UPDATE jewelry SET active = 1 FROM jewelry INNER JOIN ProductDetails ON jewelry.ProductID = ProductDetails.ProductID WHERE ProductDetailID = " & rsUpdate.Fields.Item("DetailID").Value
	objCmd.Execute()
	
	' Delete the item
	set objCmd = Server.CreateObject("ADODB.Command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "DELETE FROM TBL_OrderSummary WHERE OrderDetailID = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("detailid",3,1,15,request("detailid")))
	objCmd.Execute()
	
	
	
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
	
	
	
	' If there was a coupon used then deduct the coupon and tax amount from the order
	if rsGetOrder.Fields.Item("coupon_code").Value <> "" then
	
		if NOT rsGetCouponDiscount.eof then
			var_discount_difference = request("item_origprice") - request("item_price")
			discount = ""
			
			if rsGetOrder.Fields.Item("coupon_code").Value = "YTG89R57" then
				discount = "total_preferred_discount = total_preferred_discount - " & var_discount_difference
			else
				discount = "total_coupon_discount = total_coupon_discount - " & var_discount_difference
			end if
		end if
		

	response.write "item price: " & request("item_price") & "<br/>"
	response.write "orig price: " & request("item_origprice")  & "<br/>"	
	response.write "discount: " & discount & "<br/>"
	response.write "var_discount_difference: " & var_discount_difference & "<br/>"


		set objCmd = Server.CreateObject("ADODB.Command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "UPDATE sent_items SET " &  discount & " WHERE ID = ?" 
		objCmd.Parameters.Append(objCmd.CreateParameter("string_id",3,1,12,request("invoiceid")))
		objCmd.Execute()

	end if	' if coupon or  tax needs to be calculated

if request("item_price") <> 0 then
	item_price = request("item_price")
else
	item_price = request("item_origprice")
end if


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
			"""amount"":""" & item_price & """," & _
			"""line_items"": [{" & _
				"""id"":""1""," & _
				"""quantity"": 1," & _
				"""unit_price"": " & item_price & "," & _
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



	set objCmd = Server.CreateObject("ADODB.Command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "UPDATE sent_items SET total_sales_tax = total_sales_tax - " & amount_to_collect & ", taxes_state_only = taxes_state_only - " & state_tax_collectable & ", taxes_county_only = taxes_county_only - " & county_tax_collectable & ", taxes_city_only = taxes_city_only - " & city_tax_collectable & ", taxes_special_only = taxes_special_only - " & special_district_tax_collectable & " WHERE ID = ?" 
	objCmd.Parameters.Append(objCmd.CreateParameter("string_id",3,1,12,request("invoiceid")))
	objCmd.Execute()

end if ' if there are taxes to calc

%>