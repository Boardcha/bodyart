<%
if request.form("page_name") = "cart.asp" or request.form("page_name") = "cart.asp?usecredit=yes" or request.form("page_name") = "cart.asp?remove_coupon=yes" or Request.ServerVariables("URL") = "/cart.asp" or Request.ServerVariables("URL") = "/cart2.asp" then
DiscountSubtotal = 50

' THIS CODE SHOULD BE ALMOST IDENTICAL TO /checkout/ajax-display-shipping-usps.asp


' INTERNATIONAL -----------------------------
if strcountryName <> "US" AND strcountryName <> "CA" AND strcountryName <> "" then
	var_country = "International"
	sql_price = "ShippingAmount AS price"
	
	if CCur(var_subtotal_after_discounts) >= CCur(DiscountSubtotal) then
		sql_price = "(ShippingAmount) AS price"
	end if
	
end if

' CANADA -----------------------------
if strcountryName = "CA" then
	var_country = "Canada"
	sql_price = "ShippingAmount AS price" 

	if CCur(var_subtotal_after_discounts) >= CCur(DiscountSubtotal) then
		sql_price = "(ShippingAmount) AS price"
	end if

end if

' USA --------------------------------
if strcountryName  = "US" then
	var_shipping_goal = "FREE"
	var_country = "USA"
	sql_price = "ShippingAmount AS price"
	
	if CCur(var_subtotal_after_discounts) >= CCur(DiscountSubtotal) then
		sql_price = "(ShippingAmount - ShippingDiscount) AS price"
	end if
end if

set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT IDShipping, " & sql_price & ", ShippingName, ShippingDesc_Public, ShippingWeight, ShippingType, ShippingDiscount, DiscountSubtotal, est_days_min, est_days_max, Shipping_ActualPrice, Shipping_NameWebDisplay, FRMSelected, EstimatedShipDays FROM dbo.TBL_ShippingMethods WHERE ShippingType = ? ORDER BY price ASC"
objCmd.Parameters.Append(objCmd.CreateParameter("Country",200,1,30,var_country))
Set rsGetShippingOptions = objCmd.Execute()

if NOT rsGetShippingOptions.EOF then

	var_shipping_cost = FormatNumber(rsGetShippingOptions.Fields.Item("price").Value,2)
	session("shipping_cost") = var_shipping_cost
	
	var_shipping_cost_friendly = FormatCurrency(rsGetShippingOptions.Fields.Item("price").Value,2)
	
	if var_shipping_cost <= 0 then
		var_shipping_cost_friendly = "FREE"
	end if
 
else
	var_shipping_cost_friendly = ""
end if

if (session("weight") > 32 and session("weight") <= 200) and var_country = "USA"  then
	var_shipping_cost_friendly = "(Large or heavy items)<br/>" + replace(var_shipping_cost_friendly, "FREE", "")
End If ' weight notice

set rsGetShippingOptions = nothing


end if ' if session("cart_page") = "yes"




 %>    