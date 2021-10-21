<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001" %>
<!--#include virtual="/template/inc_includes_ajax.asp" -->
<!--#include virtual="cart/inc_cart_main.asp"-->
<!--#include virtual="cart/inc_cart_loopitems-begin.asp"-->
<!--#include virtual="cart/inc_cart_loopitems-end.asp"-->
<% 
' This ajax page updates shipping cost when user changed shipping option on the payment sheet

shipping_id = request("shipping_id")

If shipping_id <> "" Then

	DiscountSubtotal = 25

	sql_price = "ShippingAmount AS price"	
	If CCur(var_subtotal_after_discounts) >= CCur(DiscountSubtotal) Then
		sql_price = "(ShippingAmount - ShippingDiscount) AS price"
	End If
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT IDShipping, " & sql_price & " FROM dbo.TBL_ShippingMethods WHERE IDShipping = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("IDShipping",200,1,30,shipping_id))
	Set rsGetShipping = objCmd.Execute()

	If NOT rsGetShipping.EOF  Then
		%>		
		
		{"cost": "<%=rsGetShipping("price")%>"}
 
		<% 
	Else
		%>		
		
		{"cost": "0"}
 
		<% 	
	End If 
	set rsGetShipping = nothing
End If 


DataConn.Close()
Set DataConn = Nothing
%>