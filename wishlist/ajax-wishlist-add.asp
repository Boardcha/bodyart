<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/template/inc_includes_ajax.asp" -->

<%
session("last_wishlist_item") = ""

' ADD item to wishlist -------------------------------
if request.form("qty") <> "" then
	var_add_qty = request.form("qty")
else
	var_add_qty = 1
end if


	If Request.Form("WishlistCategory") <> "" then Category = Request.Form("WishlistCategory") else Category = NULL

		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "INSERT INTO wishlist (custID, itemID, itemDetailID, dateadded, desired, comments, WishlistID, priority) VALUES (?, ?, ?, ?, ?, ?, ?, ?)"
		objCmd.NamedParameters = True
		objCmd.Parameters.Append(objCmd.CreateParameter("@CustID_Cookie",3,1,10,CustID_Cookie))
		objCmd.Parameters.Append(objCmd.CreateParameter("@ItemID",3,1,10,request.form("productid")))
		objCmd.Parameters.Append(objCmd.CreateParameter("@ItemDetailID",3,1,10,request.form("add-cart")))
		objCmd.Parameters.Append(objCmd.CreateParameter("@DateAdded",200,1,12,date()))
		objCmd.Parameters.Append(objCmd.CreateParameter("@Desired",3,1,10,var_add_qty))
		
		if Request.Form("preorder_specs") <> "" then	
			wishlist_preorders = Request.Form("preorder_specs")
		else
			wishlist_preorders = ""
		end if
		
		objCmd.Parameters.Append(objCmd.CreateParameter("@Preorders",200,1,500,wishlist_preorders))
		objCmd.Parameters.Append(objCmd.CreateParameter("@Category",3,1,10, 0))
		objCmd.Parameters.Append(objCmd.CreateParameter("@Priority",200,1,50,3))
		objCmd.Execute()


		' ------------------------------------------
		' Get the last added item and store it to a session in case user wants to update the detail from the productdetails.asp page
		' ------------------------------------------
		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "SELECT ID FROM wishlist WHERE custID = ? ORDER BY ID DESC"
		objCmd.Parameters.Append(objCmd.CreateParameter("customerID",3,1,10,CustID_Cookie))
		Set rsGetLastItem = objCmd.Execute()

		if NOT rsGetLastItem.EOF then
			session("last_wishlist_item") = rsGetLastItem.Fields.Item("ID").Value
		end if
%>
		{  
			"last_wishlist_item":"<%= session("last_wishlist_item") %>"		
		}
<%

DataConn.Close()
Set DataConn = Nothing
%>
